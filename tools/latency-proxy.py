#!/usr/bin/env python3
"""
Latency-injecting TCP proxy for testing HomeFront / FlexKit on a slow link.

Blazor Server sends every keystroke and click over a WebSocket, so felt speed is
dominated by round-trip time. Chrome DevTools' network throttling does NOT apply
to WebSocket frames, which makes the app look fast there no matter what you set.
This proxy sits in front of the dev server and delays traffic in both directions,
so WebSocket frames are slowed exactly like everything else.

    # dev server already running on 5266
    python3 latency-proxy.py --listen 5267 --target 5266 --rtt 150

    then browse http://localhost:5267  (instead of 5266)

--rtt is the ROUND TRIP added, split evenly between the two directions, which is
what a real network does. Useful settings:

    --rtt 20     same-building LAN
    --rtt 60     typical broadband to a regional data centre
    --rtt 150    cross-country VPN / busy office wifi   <- where bugs show up
    --rtt 300    satellite / poor hotel wifi

Add --jitter to vary each delay randomly by up to that many ms, which is what
makes a bad link feel "sticky" rather than merely slow.

No dependencies, no sudo, no system settings touched. Ctrl-C to stop.
"""

import argparse
import asyncio
import random
import time

parser = argparse.ArgumentParser(description="Delay TCP traffic to simulate a slow network.")
parser.add_argument("--listen", type=int, default=5267, help="port to browse (default 5267)")
parser.add_argument("--target", type=int, default=5266, help="port the dev server is on (default 5266)")
parser.add_argument("--host", default="127.0.0.1", help="dev server host (default 127.0.0.1)")
parser.add_argument("--rtt", type=float, default=150.0, help="round-trip ms to add (default 150)")
parser.add_argument("--jitter", type=float, default=0.0, help="max extra random ms per hop (default 0)")
parser.add_argument("--quiet", action="store_true", help="do not log connections")
args = parser.parse_args()

ONE_WAY = args.rtt / 2000.0          # ms round trip -> seconds one way
JITTER = args.jitter / 1000.0
_conns = 0


async def pump(reader, writer, label):
    """Copy reader->writer, holding every chunk for the one-way delay.

    Chunks are timestamped on arrival and released on schedule, so a burst is
    delayed rather than serialized — the same shape a real link produces.
    """
    queue = asyncio.Queue()

    async def release():
        try:
            while True:
                item = await queue.get()
                if item is None:
                    break
                due, chunk = item
                now = time.monotonic()
                if due > now:
                    await asyncio.sleep(due - now)
                writer.write(chunk)
                await writer.drain()
        except (ConnectionResetError, BrokenPipeError):
            pass
        finally:
            try:
                writer.close()
            except Exception:
                pass

    releaser = asyncio.create_task(release())
    try:
        while True:
            chunk = await reader.read(65536)
            if not chunk:
                break
            delay = ONE_WAY + (random.random() * JITTER if JITTER else 0.0)
            await queue.put((time.monotonic() + delay, chunk))
    except (ConnectionResetError, BrokenPipeError):
        pass
    finally:
        await queue.put(None)
        await releaser


async def handle(client_reader, client_writer):
    global _conns
    _conns += 1
    n = _conns
    if not args.quiet:
        print(f"  [{n}] connect")
    try:
        server_reader, server_writer = await asyncio.open_connection(args.host, args.target)
    except OSError as exc:
        print(f"  [{n}] cannot reach {args.host}:{args.target} — is the dev server running? ({exc})")
        client_writer.close()
        return

    await asyncio.gather(
        pump(client_reader, server_writer, "up"),
        pump(server_reader, client_writer, "down"),
        return_exceptions=True,
    )
    if not args.quiet:
        print(f"  [{n}] close")


async def main():
    server = await asyncio.start_server(handle, "127.0.0.1", args.listen)
    print(f"Adding {args.rtt:.0f} ms round trip"
          + (f" (+ up to {args.jitter:.0f} ms jitter per hop)" if args.jitter else "")
          + f" to 127.0.0.1:{args.target}")
    print(f"Browse http://localhost:{args.listen}   (Ctrl-C to stop)")
    async with server:
        await server.serve_forever()


try:
    asyncio.run(main())
except KeyboardInterrupt:
    print("\nstopped")
