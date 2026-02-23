"""Full end-to-end MCP server test."""
import asyncio, json, sys, time

PYTHON = r"C:\Users\12152\AppData\Local\Programs\Python\Python313\python.exe"
PASS = "PASS"
FAIL = "FAIL"

async def read_response(stdout, expected_id, timeout=120):
    while True:
        line = await asyncio.wait_for(stdout.readline(), timeout=timeout)
        if not line:
            raise EOFError("Server closed")
        text = line.decode().strip()
        if not text:
            continue
        try:
            msg = json.loads(text)
        except json.JSONDecodeError:
            continue
        if "id" in msg and msg["id"] == expected_id:
            return msg

async def send(proc, msg):
    data = json.dumps(msg) + "\n"
    proc.stdin.write(data.encode("utf-8"))
    await proc.stdin.drain()

async def main():
    results = []
    proc = await asyncio.create_subprocess_exec(
        PYTHON, "outlook_mcp.py",
        stdin=asyncio.subprocess.PIPE,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
        limit=50 * 1024 * 1024,
    )
    await asyncio.sleep(1)

    # ---- Test 1: Initialize ----
    t0 = time.time()
    try:
        await send(proc, {
            "jsonrpc": "2.0", "id": 1, "method": "initialize",
            "params": {"protocolVersion": "2024-11-05", "capabilities": {},
                       "clientInfo": {"name": "e2e-test", "version": "1.0"}},
        })
        resp = await read_response(proc.stdout, 1)
        await send(proc, {"jsonrpc": "2.0", "method": "notifications/initialized", "params": {}})
        assert "result" in resp
        assert "capabilities" in resp["result"]
        server_name = resp["result"].get("serverInfo", {}).get("name", "?")
        results.append((PASS, "Initialize", f"server={server_name}", time.time() - t0))
    except Exception as e:
        results.append((FAIL, "Initialize", str(e), time.time() - t0))

    # ---- Test 2: List Tools ----
    t0 = time.time()
    try:
        await send(proc, {"jsonrpc": "2.0", "id": 2, "method": "tools/list", "params": {}})
        resp = await read_response(proc.stdout, 2)
        tools = resp["result"]["tools"]
        tool_names = [t["name"] for t in tools]
        assert len(tools) >= 3, f"Expected >=3 tools, got {len(tools)}"
        assert "check_mailbox_access" in tool_names
        assert "get_email_chain" in tool_names
        assert "get_email_contacts" in tool_names
        results.append((PASS, "List Tools", f"{len(tools)} tools: {tool_names}", time.time() - t0))
    except Exception as e:
        results.append((FAIL, "List Tools", str(e), time.time() - t0))

    # ---- Test 3: check_mailbox_access ----
    t0 = time.time()
    try:
        await send(proc, {
            "jsonrpc": "2.0", "id": 3, "method": "tools/call",
            "params": {"name": "check_mailbox_access", "arguments": {}},
        })
        resp = await read_response(proc.stdout, 3)
        text = resp["result"]["content"][0]["text"]
        try:
            d = json.loads(text)
        except json.JSONDecodeError:
            import ast; d = ast.literal_eval(text)
        status = d.get("status", "?")
        personal = d.get("personal_mailbox", {}).get("accessible", False)
        shared_cfg = d.get("shared_mailbox", {}).get("configured", False)
        results.append((PASS if status == "success" else FAIL, "check_mailbox_access",
                        f"status={status}, personal={personal}, shared_configured={shared_cfg}",
                        time.time() - t0))
    except Exception as e:
        results.append((FAIL, "check_mailbox_access", str(e), time.time() - t0))

    # ---- Test 4: get_email_chain("Eleven") ----
    t0 = time.time()
    try:
        await send(proc, {
            "jsonrpc": "2.0", "id": 4, "method": "tools/call",
            "params": {"name": "get_email_chain",
                       "arguments": {"search_text": "Eleven", "include_shared": False}},
        })
        resp = await read_response(proc.stdout, 4)
        text = resp["result"]["content"][0]["text"]
        try:
            d = json.loads(text)
        except json.JSONDecodeError:
            import ast; d = ast.literal_eval(text)
        status = d.get("status", "?")
        total = d.get("summary", {}).get("total_emails", 0)
        convos = d.get("summary", {}).get("conversations", 0)
        results.append((PASS if status == "success" and total > 0 else FAIL,
                        'get_email_chain("Eleven")',
                        f"status={status}, emails={total}, conversations={convos}",
                        time.time() - t0))
    except Exception as e:
        results.append((FAIL, 'get_email_chain("Eleven")', str(e), time.time() - t0))

    # ---- Test 5: get_email_contacts("Eleven") ----
    t0 = time.time()
    try:
        await send(proc, {
            "jsonrpc": "2.0", "id": 5, "method": "tools/call",
            "params": {"name": "get_email_contacts",
                       "arguments": {"search_text": "Eleven", "include_shared": False}},
        })
        resp = await read_response(proc.stdout, 5)
        text = resp["result"]["content"][0]["text"]
        try:
            d = json.loads(text)
        except json.JSONDecodeError:
            import ast; d = ast.literal_eval(text)
        status = d.get("status", "?")
        contacts = d.get("contacts", [])
        results.append((PASS if status == "success" and len(contacts) > 0 else FAIL,
                        'get_email_contacts("Eleven")',
                        f"status={status}, contacts={len(contacts)}",
                        time.time() - t0))
    except Exception as e:
        results.append((FAIL, 'get_email_contacts("Eleven")', str(e), time.time() - t0))

    # ---- Test 6: get_email_chain with empty search (edge case) ----
    t0 = time.time()
    try:
        await send(proc, {
            "jsonrpc": "2.0", "id": 6, "method": "tools/call",
            "params": {"name": "get_email_chain",
                       "arguments": {"search_text": "xyzzy_no_match_12345"}},
        })
        resp = await read_response(proc.stdout, 6)
        text = resp["result"]["content"][0]["text"]
        try:
            d = json.loads(text)
        except json.JSONDecodeError:
            import ast; d = ast.literal_eval(text)
        status = d.get("status", "?")
        results.append((PASS if status == "no_emails_found" else FAIL,
                        'get_email_chain(no match)',
                        f"status={status}",
                        time.time() - t0))
    except Exception as e:
        results.append((FAIL, 'get_email_chain(no match)', str(e), time.time() - t0))

    # ---- Cleanup ----
    proc.stdin.close()
    try:
        await asyncio.wait_for(proc.wait(), timeout=5)
    except asyncio.TimeoutError:
        proc.kill()

    # ---- Report ----
    print("\n" + "=" * 65)
    print("  MCP Server End-to-End Test Results")
    print("=" * 65)
    passed = sum(1 for r in results if r[0] == PASS)
    total = len(results)
    for status, name, detail, elapsed in results:
        icon = "+" if status == PASS else "X"
        print(f"  [{icon}] {name} ({elapsed:.1f}s)")
        print(f"      {detail}")
    print("-" * 65)
    print(f"  {passed}/{total} passed")
    print("=" * 65)

asyncio.run(main())
