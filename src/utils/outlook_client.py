"""High-performance Outlook client for mailbox access and email search."""

import win32com.client
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional, Tuple
import logging
import pythoncom
import re
import time
import threading
from functools import lru_cache
from concurrent.futures import ThreadPoolExecutor, as_completed
import queue

from ..config.config_reader import config

logger = logging.getLogger(__name__)

# Outlook Recipient types
OL_TO = 1
OL_CC = 2
OL_BCC = 3

# MAPI property tag for SMTP address (resolves Exchange DNs to real emails)
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"


class OutlookClient:
    """High-performance client for accessing Outlook mailboxes with optimized search."""

    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.connected = False
        self._search_cache = {}  # Cache for search results
        self._folder_cache = {}  # Cache for folder references
        self._shared_recipient_cache = None  # Cache for resolved shared recipient
        self._connection_retry_count = 0
        self._max_retries = config.get_int('max_connection_retries', 3)

    def connect(self, retry_attempt: int = 0) -> bool:
        """Connect to Outlook application with retry logic.

        IMPORTANT: This must be called from a thread where COM is already
        initialized (e.g. the dedicated COM thread).  COM apartment setup
        is managed by outlook_mcp._com_executor, not here.
        """
        try:
            logger.info("Connecting to Outlook...")
            start_time = time.time()

            # Try to connect to existing Outlook instance first (much faster)
            try:
                self.outlook = win32com.client.GetActiveObject("Outlook.Application")
                logger.info("Connected to existing Outlook instance")
            except Exception as e:
                # GetActiveObject fails if no running instance; fall back to Dispatch
                logger.info(f"No active Outlook instance ({type(e).__name__}), creating via Dispatch...")
                self.outlook = win32com.client.Dispatch("Outlook.Application")

            self.namespace = self.outlook.GetNamespace("MAPI")

            # Try Extended MAPI login to reduce security prompts (if enabled)
            if config.get_bool('use_extended_mapi_login', True):
                try:
                    self.namespace.Logon(None, None, False, True)
                    logger.info("Extended MAPI login successful")
                except Exception as logon_error:
                    logger.warning(f"Extended MAPI login failed (non-fatal): {logon_error}")

            self.connected = True
            connection_time = time.time() - start_time
            logger.info(f"Connected to Outlook in {connection_time:.2f}s")
            return True
        except Exception as e:
            logger.error(f"Failed to connect to Outlook (attempt {retry_attempt + 1}): {e}")
            self.connected = False
            self.outlook = None
            self.namespace = None

            # Exponential backoff retry
            if retry_attempt < self._max_retries - 1:
                wait_time = 2 ** retry_attempt  # 1s, 2s, 4s
                logger.info(f"Retrying connection in {wait_time}s...")
                time.sleep(wait_time)
                return self.connect(retry_attempt + 1)

            return False

    def _is_connection_alive(self) -> bool:
        """Check if the COM connection to Outlook is still valid."""
        if not self.connected or not self.namespace:
            return False
        try:
            # Lightweight COM call to verify the connection is alive
            _ = self.namespace.CurrentUser.Name
            return True
        except Exception:
            return False

    def _ensure_connected(self) -> bool:
        """Ensure a live COM connection, reconnecting if stale."""
        if self._is_connection_alive():
            return True
        # Connection is stale or missing — reset and reconnect
        logger.warning("COM connection lost or stale, reconnecting...")
        self.connected = False
        self.outlook = None
        self.namespace = None
        self._shared_recipient_cache = None
        self._folder_cache.clear()
        self._search_cache.clear()
        return self.connect()
    
    def check_access(self) -> Dict[str, Any]:
        """Check access to personal and shared mailboxes."""
        if not self._ensure_connected():
            return {"error": "Could not connect to Outlook"}
        
        result = {
            "outlook_connected": True,
            "personal_accessible": False,
            "shared_accessible": False,
            "shared_configured": bool(config.get('shared_mailbox_email')),
            "retention_personal_months": config.get_int('personal_retention_months', 6),
            "retention_shared_months": config.get_int('shared_retention_months', 12),
            "errors": []
        }
        
        # Test personal mailbox
        try:
            personal_inbox = self.namespace.GetDefaultFolder(6)  # 6 = Inbox
            if personal_inbox:
                result["personal_accessible"] = True
                result["personal_name"] = self._get_store_display_name(personal_inbox)
        except Exception as e:
            result["errors"].append(f"Personal mailbox error: {str(e)}")
        
        # Test shared mailbox
        shared_email = config.get('shared_mailbox_email')
        if shared_email:
            try:
                # Use cached recipient if available
                if not self._shared_recipient_cache:
                    self._shared_recipient_cache = self.namespace.CreateRecipient(shared_email)
                    self._shared_recipient_cache.Resolve()
                
                if self._shared_recipient_cache.Resolved:
                    shared_inbox = self.namespace.GetSharedDefaultFolder(self._shared_recipient_cache, 6)
                    if shared_inbox:
                        result["shared_accessible"] = True
                        result["shared_name"] = self._get_store_display_name(shared_inbox)
            except Exception as e:
                result["errors"].append(f"Shared mailbox error: {str(e)}")
                self._shared_recipient_cache = None  # Clear cache on error
        
        return result
    
    @staticmethod
    def _shared_mailbox_configured() -> bool:
        """Return True only if a real shared mailbox email is set (not the placeholder)."""
        addr = (config.get('shared_mailbox_email') or '').strip()
        return bool(addr) and 'example.com' not in addr and 'your-shared' not in addr

    def search_emails(self, search_text: str,
                     include_personal: bool = True,
                     include_shared: bool = True) -> List[Dict[str, Any]]:
        """Search emails in both subject and body.

        All COM work runs on the dedicated COM thread (via _run_com in
        outlook_mcp.py) to avoid cross-thread STA apartment violations.
        """
        if not self._ensure_connected():
            return []

        max_results = config.get_int('max_search_results', 500)
        cache_key = f"{search_text}_{include_personal}_{include_shared}_{max_results}"

        if cache_key in self._search_cache:
            cache_entry = self._search_cache[cache_key]
            if time.time() - cache_entry['timestamp'] < 3600:  # 1 hour cache
                logger.info(f"Returning cached results for '{search_text}'")
                return cache_entry['data']

        all_emails = []

        # --- Personal mailbox (sequential, same thread) --------------------
        if include_personal:
            personal_emails = self._search_mailbox_comprehensive(
                self.namespace.GetDefaultFolder(6),
                search_text,
                'personal',
                max_results
            )
            all_emails.extend(personal_emails)
            logger.info(f"Found {len(personal_emails)} emails in personal mailbox")

        # --- Shared mailbox (only if genuinely configured) -----------------
        if include_shared and self._shared_mailbox_configured():
            try:
                if not self._shared_recipient_cache:
                    shared_email = config.get('shared_mailbox_email')
                    self._shared_recipient_cache = self.namespace.CreateRecipient(shared_email)
                    self._shared_recipient_cache.Resolve()

                if self._shared_recipient_cache.Resolved:
                    shared_inbox = self.namespace.GetSharedDefaultFolder(
                        self._shared_recipient_cache, 6)
                    shared_emails = self._search_mailbox_comprehensive(
                        shared_inbox,
                        search_text,
                        'shared',
                        max_results - len(all_emails)
                    )
                    all_emails.extend(shared_emails)
                    logger.info(f"Found {len(shared_emails)} emails in shared mailbox")
            except Exception as e:
                logger.error(f"Error searching shared mailbox: {e}")
                self._shared_recipient_cache = None

        # Sort by received time (newest first)
        all_emails.sort(key=lambda x: x.get('received_time', datetime.min), reverse=True)
        
        # Cache results with timestamp
        limited_results = all_emails[:max_results]
        self._search_cache[cache_key] = {
            'data': limited_results,
            'timestamp': time.time()
        }
        
        # Limit cache size
        if len(self._search_cache) > 100:
            # Remove oldest entries
            oldest_key = min(self._search_cache.keys(), 
                           key=lambda k: self._search_cache[k].get('timestamp', 0))
            del self._search_cache[oldest_key]
        
        return limited_results
    
    def search_emails_by_subject(self, subject: str,
                                include_personal: bool = True,
                                include_shared: bool = True) -> List[Dict[str, Any]]:
        """Legacy method - redirects to search_emails for backward compatibility."""
        return self.search_emails(subject, include_personal, include_shared)

    def _search_mailbox_comprehensive(self, inbox_folder, search_text: str,
                                      mailbox_type: str, max_results: int) -> List[Dict[str, Any]]:
        """Search Inbox + all subfolders + Sent Items using Items.Restrict with DASL filters.

        AdvancedSearch is event-driven and requires a Windows message pump to fire
        the SearchComplete callback.  Without PumpWaitingMessages the completion
        flag never flips, so every search silently returns 0 results.
        Items.Restrict is synchronous and works reliably from any thread.
        """
        emails = []
        found_ids = set()  # Track found emails to avoid duplicates

        # --- 1. Collect every folder we need to search ----------------------
        folders_to_search = []
        # Inbox itself + all recursive subfolders
        self._collect_subfolders_recursive(inbox_folder, folders_to_search)

        # Sent Items (sits next to Inbox, under the store root)
        if config.get_bool('include_sent_items', True):
            sent = self._get_folder_by_name(inbox_folder.Parent, 'Sent Items')
            if sent:
                folders_to_search.append(sent)

        logger.info(
            f"Searching {len(folders_to_search)} folders in {mailbox_type} mailbox "
            f"for '{search_text}'"
        )

        # --- 2. Search each folder with Restrict ---------------------------
        for folder in folders_to_search:
            if len(emails) >= max_results:
                break
            try:
                folder_emails = self._search_folder_restrict(
                    folder, search_text, mailbox_type,
                    max_results - len(emails), found_ids
                )
                emails.extend(folder_emails)
            except Exception as e:
                logger.debug(f"Error searching folder {getattr(folder, 'Name', '?')}: {e}")

        return emails

    # -- Recursive subfolder collection ------------------------------------

    def _collect_subfolders_recursive(self, folder, folder_list: list):
        """Add *folder* and every subfolder beneath it to *folder_list*."""
        folder_list.append(folder)
        try:
            for i in range(1, folder.Folders.Count + 1):
                subfolder = folder.Folders.Item(i)
                self._collect_subfolders_recursive(subfolder, folder_list)
        except Exception as e:
            logger.debug(f"Error enumerating subfolders of {getattr(folder, 'Name', '?')}: {e}")

    # -- Per-folder Restrict search ----------------------------------------

    def _search_folder_restrict(self, folder, search_text: str, mailbox_type: str,
                                max_results: int, found_ids: set) -> List[Dict[str, Any]]:
        """Search a single folder using Items.Restrict with a DASL LIKE filter.

        Searches both subject and plain-text body (textdescription).
        """
        emails = []
        # Escape single-quotes for DASL string literals
        safe = search_text.replace("'", "''")

        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)  # newest first

            # DASL filter: subject OR body contains the search term (case-insensitive)
            dasl = (
                f"@SQL=(\"urn:schemas:httpmail:subject\" LIKE '%{safe}%'"
                f" OR \"urn:schemas:httpmail:textdescription\" LIKE '%{safe}%')"
            )

            filtered = items.Restrict(dasl)

            count = 0
            for item in filtered:
                if count >= max_results:
                    break
                entry_id = getattr(item, 'EntryID', '')
                if entry_id and entry_id not in found_ids:
                    email_data = self._extract_email_data(item, folder.Name, mailbox_type)
                    if email_data:
                        emails.append(email_data)
                        found_ids.add(entry_id)
                        count += 1

            logger.info(f"  {folder.Name}: {count} matches")
        except Exception as e:
            logger.warning(f"Restrict search failed in '{getattr(folder, 'Name', '?')}': {e}")
            # Last-resort fallback: subject-only filter (always works)
            try:
                subject_filter = f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{safe}%'"
                filtered = folder.Items
                filtered.Sort("[ReceivedTime]", True)
                filtered = filtered.Restrict(subject_filter)
                count = 0
                for item in filtered:
                    if count >= max_results:
                        break
                    entry_id = getattr(item, 'EntryID', '')
                    if entry_id and entry_id not in found_ids:
                        email_data = self._extract_email_data(item, folder.Name, mailbox_type)
                        if email_data:
                            emails.append(email_data)
                            found_ids.add(entry_id)
                            count += 1
                logger.info(f"  {folder.Name} (subject-only fallback): {count} matches")
            except Exception as e2:
                logger.debug(f"Subject-only fallback also failed for {getattr(folder, 'Name', '?')}: {e2}")

        return emails
    
    def _extract_email_data(self, item, folder_name: str,
                           mailbox_type: str) -> Dict[str, Any]:
        """Extract email data with full To/CC recipient lists and SMTP resolution."""
        try:
            # Get the full email body
            body = getattr(item, 'Body', '')

            # Apply max_body_chars if configured (0 means no limit)
            max_body_chars = config.get_int('max_body_chars', 0)
            if max_body_chars > 0 and len(body) > max_body_chars:
                body = body[:max_body_chars] + " [truncated]"

            # Clean HTML if configured
            if config.get_bool('clean_html_content', True) and body:
                body = self._clean_html(body)

            # --- Recipient extraction with To / CC split --------------------
            to_recipients = []
            cc_recipients = []
            all_recipients = []   # backwards-compat flat list
            max_recipients = config.get_int('max_recipients_display', 50)

            try:
                count = 0
                for recipient in item.Recipients:
                    if count >= max_recipients:
                        remaining = item.Recipients.Count - count
                        overflow = f"... and {remaining} more"
                        all_recipients.append(overflow)
                        break

                    name = getattr(recipient, 'Name', '')
                    address = getattr(recipient, 'Address', '')

                    # Resolve Exchange DN to SMTP address
                    smtp = address
                    if address and (address.startswith('/') or '@' not in address):
                        try:
                            smtp = recipient.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
                        except Exception:
                            smtp = address  # keep the raw value

                    entry = {'name': name, 'email': smtp}
                    rtype = getattr(recipient, 'Type', OL_TO)

                    if rtype == OL_CC:
                        cc_recipients.append(entry)
                    else:  # OL_TO (and BCC lumped with TO for simplicity)
                        to_recipients.append(entry)

                    all_recipients.append(name or smtp)
                    count += 1
            except Exception:
                pass

            # Resolve sender SMTP address too
            sender_email = getattr(item, 'SenderEmailAddress', '')
            if sender_email and (sender_email.startswith('/') or '@' not in sender_email):
                try:
                    sender = item.Sender
                    if sender:
                        sender_email = sender.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
                except Exception:
                    # Fallback: try the mail item's own property accessor
                    try:
                        sender_email = item.PropertyAccessor.GetProperty(
                            "http://schemas.microsoft.com/mapi/proptag/0x5D01001E"
                        )  # PR_SENDER_SMTP_ADDRESS
                    except Exception:
                        pass  # keep original Exchange DN

            email_data = {
                'subject': getattr(item, 'Subject', 'No Subject'),
                'sender_name': getattr(item, 'SenderName', 'Unknown'),
                'sender_email': sender_email,
                'to_recipients': to_recipients,
                'cc_recipients': cc_recipients,
                'recipients': all_recipients,  # flat list for backwards compat
                'received_time': getattr(item, 'ReceivedTime', datetime.now()),
                'folder_name': folder_name,
                'mailbox_type': mailbox_type,
                'importance': getattr(item, 'Importance', 1),
                'body': body,
                'size': getattr(item, 'Size', 0),
                'attachments_count': getattr(item.Attachments, 'Count', 0) if hasattr(item, 'Attachments') else 0,
                'unread': getattr(item, 'Unread', False),
                'entry_id': getattr(item, 'EntryID', '')
            }

            # Release COM reference to free memory
            item = None

            return email_data
        except Exception as e:
            logger.error(f"Error extracting email data: {e}")
            return None
    
    def _get_store_display_name(self, folder) -> str:
        """Safely get store display name from a folder."""
        try:
            if hasattr(folder, 'Parent'):
                parent = folder.Parent
                if hasattr(parent, 'DisplayName'):
                    return parent.DisplayName
                elif hasattr(parent, 'Name'):
                    return parent.Name
            return "Mailbox"
        except:
            return "Mailbox"
    
    def _get_folder_by_name(self, parent, name: str):
        """Get a child folder by name.  Accepts either a Store or a MAPIFolder."""
        cache_key = f"{id(parent)}_{name}"

        if cache_key in self._folder_cache:
            return self._folder_cache[cache_key]

        try:
            # If it's a Store, drill into the root folder first
            folders = (parent.GetRootFolder().Folders
                       if hasattr(parent, 'GetRootFolder')
                       else parent.Folders)
            for folder in folders:
                if folder.Name.lower() == name.lower():
                    self._folder_cache[cache_key] = folder
                    return folder
        except Exception:
            pass

        return None
    
    def _clean_html(self, text: str) -> str:
        """Clean HTML from email body."""
        import re
        
        # Remove HTML tags
        text = re.sub(r'<[^>]+>', '', text)
        
        # Decode common HTML entities
        html_entities = {
            '&amp;': '&',
            '&lt;': '<',
            '&gt;': '>',
            '&quot;': '"',
            '&#39;': "'",
            '&nbsp;': ' '
        }
        
        for entity, char in html_entities.items():
            text = text.replace(entity, char)
        
        # Clean up whitespace
        text = re.sub(r'\s+', ' ', text).strip()
        
        return text


# Global client instance
outlook_client = OutlookClient()
