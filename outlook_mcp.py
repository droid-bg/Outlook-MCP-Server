"""Simplified Outlook MCP Server with three main tools."""

import asyncio
import atexit
import logging
import platform
import sys
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Sequence

# Check if running on Windows
if platform.system() != 'Windows':
    print("[ERROR] Outlook MCP Server requires Windows with Microsoft Outlook installed")
    print(f"   Current platform: {platform.system()}")
    print("\n[INFO] To use this server:")
    print("   1. Run on a Windows machine with Outlook installed")
    print("   2. Or use a Windows virtual machine")
    print("   3. Or access a remote Windows desktop")
    sys.exit(1)

import pythoncom
from mcp import server, types
from mcp.server import Server
from mcp.server.stdio import stdio_server

try:
    from src.config.config_reader import config
    from src.utils.outlook_client import outlook_client
    from src.utils.email_formatter import format_mailbox_status, format_email_chain
except ImportError as e:
    print(f"[ERROR] Import Error: {e}")
    print("\n[INFO] Please install required dependencies:")
    print("   pip install -r requirements.txt")
    print("\nNote: pywin32 is required and only works on Windows")
    sys.exit(1)

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Dedicated COM thread
#
# Outlook COM objects live in an STA (Single-Threaded Apartment) and MUST be
# accessed from the same thread that created them.  asyncio.to_thread() uses
# a general thread pool where successive calls can land on different threads,
# causing "RPC server is unavailable" errors.
#
# We create a single-worker ThreadPoolExecutor so every COM call runs on the
# same thread, and initialise COM once on that thread at startup.
# ---------------------------------------------------------------------------
_com_executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="outlook-com")


def _init_com_thread():
    """Called once on the dedicated thread to set up the COM apartment."""
    pythoncom.CoInitialize()


# Block until COM is initialised on the worker thread
_com_executor.submit(_init_com_thread).result()


def _shutdown_com():
    """Clean up the COM thread on interpreter exit."""
    try:
        _com_executor.submit(pythoncom.CoUninitialize).result(timeout=5)
    except Exception:
        pass
    _com_executor.shutdown(wait=False)


atexit.register(_shutdown_com)


async def _run_com(func, *args, **kwargs):
    """Run *func* on the dedicated COM thread (non-blocking to the event loop)."""
    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(
        _com_executor, lambda: func(*args, **kwargs)
    )

# Create MCP server
app = Server("outlook-mcp-server")


@app.list_tools()
async def list_tools() -> list[types.Tool]:
    """List available MCP tools."""
    return [
        types.Tool(
            name="check_mailbox_access",
            description="Check connection status and access to personal and shared mailboxes with retention policy info",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": []
            }
        ),
        types.Tool(
            name="get_email_chain",
            description="Searches for emails containing the specified text in BOTH subject and body. Searches Inbox, all Inbox subfolders, and Sent Items. Returns full email content including sender, To/CC recipients with email addresses, timestamps, and message bodies.",
            inputSchema={
                "type": "object",
                "properties": {
                    "search_text": {
                        "type": "string",
                        "description": "Text to search for in email subject and body (case-insensitive substring match)."
                    },
                    "include_personal": {
                        "type": "boolean",
                        "description": "Search personal mailbox (default: true)",
                        "default": True
                    },
                    "include_shared": {
                        "type": "boolean",
                        "description": "Search shared mailbox (default: true)",
                        "default": True
                    }
                },
                "required": ["search_text"]
            }
        ),
        types.Tool(
            name="get_email_contacts",
            description="Contact intelligence tool. Searches emails by keyword, then returns a ranked list of every person on those threads — with name, SMTP email, and how many times they appeared as sender, To, or CC. Use this to map who is involved in order threads, vendor conversations, or project discussions.",
            inputSchema={
                "type": "object",
                "properties": {
                    "search_text": {
                        "type": "string",
                        "description": "Keyword to search for (e.g. a vendor name, order number, project)."
                    },
                    "include_personal": {
                        "type": "boolean",
                        "description": "Search personal mailbox (default: true)",
                        "default": True
                    },
                    "include_shared": {
                        "type": "boolean",
                        "description": "Search shared mailbox (default: true)",
                        "default": True
                    }
                },
                "required": ["search_text"]
            }
        )
    ]


@app.call_tool()
async def call_tool(name: str, arguments: dict[str, Any]) -> Sequence[types.TextContent]:
    """Handle tool calls."""
    
    logger.info(f"Executing tool: {name}")
    
    try:
        if name == "check_mailbox_access":
            return await handle_check_mailbox_access()
            
        elif name == "get_email_chain":
            search_text = arguments.get("search_text")
            if not search_text:
                raise ValueError("search_text parameter is required")

            include_personal = arguments.get("include_personal", True)
            include_shared = arguments.get("include_shared", True)

            return await handle_get_email_chain(search_text, include_personal, include_shared)

        elif name == "get_email_contacts":
            search_text = arguments.get("search_text")
            if not search_text:
                raise ValueError("search_text parameter is required")

            include_personal = arguments.get("include_personal", True)
            include_shared = arguments.get("include_shared", True)

            return await handle_get_email_contacts(search_text, include_personal, include_shared)

        else:
            raise ValueError(f"Unknown tool: {name}")
            
    except Exception as e:
        logger.error(f"Error in tool {name}: {e}")
        error_response = {
            "status": "error",
            "tool": name,
            "error": str(e),
            "message": f"Failed to execute {name}: {str(e)}"
        }
        return [types.TextContent(type="text", text=str(error_response))]


async def handle_check_mailbox_access():
    """Handle mailbox access check."""
    logger.info("Checking mailbox access...")
    
    try:
        # Check access to mailboxes (runs on dedicated COM thread)
        access_result = await _run_com(outlook_client.check_access)
        
        # Format response
        formatted_result = format_mailbox_status(access_result)
        
        logger.info("Mailbox access check completed")
        return [types.TextContent(type="text", text=str(formatted_result))]
        
    except Exception as e:
        logger.error(f"Error checking mailbox access: {e}")
        error_response = {
            "status": "error",
            "message": f"Could not check mailbox access: {str(e)}",
            "troubleshooting": [
                "Make sure Outlook is running",
                "Grant permission when security dialog appears", 
                "Check network connectivity"
            ]
        }
        return [types.TextContent(type="text", text=str(error_response))]


async def handle_get_email_chain(search_text: str, include_personal: bool, include_shared: bool):
    """Handle email search and retrieval."""
    logger.info(f"Searching for emails containing: {search_text}")
    
    try:
        # Search for emails in both subject and body (runs on dedicated COM thread)
        emails = await _run_com(
            outlook_client.search_emails,
            search_text=search_text,
            include_personal=include_personal,
            include_shared=include_shared
        )
        
        # Format response
        formatted_result = format_email_chain(emails, search_text)
        
        logger.info(f"Found {len(emails)} emails containing '{search_text}'")
        return [types.TextContent(type="text", text=str(formatted_result))]
        
    except Exception as e:
        logger.error(f"Error searching emails: {e}")
        error_response = {
            "status": "error", 
            "search_text": search_text,
            "message": f"Could not search emails: {str(e)}",
            "troubleshooting": [
                "Verify Outlook connection", 
                "Use specific search terms for best results",
                "Ensure mailboxes are accessible"
            ]
        }
        return [types.TextContent(type="text", text=str(error_response))]


async def handle_get_email_contacts(search_text: str, include_personal: bool, include_shared: bool):
    """Search emails, then return a ranked contact/participant list."""
    logger.info(f"Contact intelligence search for: {search_text}")

    try:
        emails = await _run_com(
            outlook_client.search_emails,
            search_text=search_text,
            include_personal=include_personal,
            include_shared=include_shared
        )

        # Build the participant list via the formatter helper
        from src.utils.email_formatter import get_participants, get_date_range
        participants = get_participants(emails)

        result = {
            "status": "success" if emails else "no_emails_found",
            "search_text": search_text,
            "emails_scanned": len(emails),
            "date_range": get_date_range(emails),
            "contacts": participants,
        }

        logger.info(f"Found {len(participants)} unique contacts across {len(emails)} emails")
        return [types.TextContent(type="text", text=str(result))]

    except Exception as e:
        logger.error(f"Error in contact search: {e}")
        return [types.TextContent(type="text", text=str({
            "status": "error",
            "search_text": search_text,
            "message": f"Contact search failed: {e}"
        }))]


@app.list_resources()
async def list_resources() -> list[types.Resource]:
    """List available resources."""
    return [
        types.Resource(
            uri="outlook-mcp://config",
            name="Current Configuration", 
            description="Show current configuration settings",
            mimeType="text/plain"
        )
    ]


@app.read_resource()
async def read_resource(uri: str) -> str:
    """Read resource content."""
    if uri == "outlook-mcp://config":
        config.show_config()
        return "Configuration displayed in console"
    else:
        raise ValueError(f"Unknown resource: {uri}")


async def main():
    """Main entry point."""
    print("=" * 60)
    print("[STARTING] Outlook MCP Server")
    print("=" * 60)
    
    # Show configuration
    config.show_config()
    
    # Important notes
    print("\n[INFO] Important Notes:")
    print("   * Make sure Microsoft Outlook is running")
    print("   * Grant permission when security dialog appears")  
    print("   * Update config.properties with your shared mailbox details")
    print("   * Server searches ALL folders, not just Inbox")
    
    shared_email = config.get('shared_mailbox_email')
    if not shared_email or 'your-shared-mailbox' in shared_email or 'example.com' in shared_email:
        print("\n[WARNING] Shared mailbox not configured!")
        print("   Update 'shared_mailbox_email' in config.properties")
    
    print("\n[TOOLS] Available Tools:")
    print("   1. check_mailbox_access - Test connection and access")
    print("   2. get_email_chain - Search emails by text in subject AND body")
    print("   3. get_email_contacts - Contact intelligence (who is on which threads)")
    
    print(f"\n[READY] Server ready! Listening for MCP client connections...")
    print("=" * 60)
    
    # Start server
    async with stdio_server() as (read_stream, write_stream):
        await app.run(read_stream, write_stream, app.create_initialization_options())


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n[INFO] Server stopped by user")
    except Exception as e:
        print(f"\n[ERROR] Server error: {e}")
        logger.error(f"Server error: {e}")
