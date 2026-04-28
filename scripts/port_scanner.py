#!/usr/bin/env python3
"""
Port Scanner - TCP Connect Scan
Works on iSH (iOS) and any Linux system.
Use only on authorized targets (e.g. scanme.nmap.org, TryHackMe machines).
"""

import socket
import sys
import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

COMMON_PORTS = [
    21, 22, 23, 25, 53, 80, 110, 111, 135, 139,
    143, 443, 445, 993, 995, 1723, 3306, 3389,
    5900, 8080, 8443, 8888
]

def grab_banner(ip, port, timeout=1):
    try:
        s = socket.socket()
        s.settimeout(timeout)
        s.connect((ip, port))
        s.send(b"HEAD / HTTP/1.0\r\n\r\n")
        banner = s.recv(1024).decode(errors="ignore").strip()
        s.close()
        return banner.split("\n")[0] if banner else ""
    except:
        return ""

def scan_port(ip, port, timeout=1, banner=False):
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(timeout)
        result = s.connect_ex((ip, port))
        s.close()
        if result == 0:
            b = grab_banner(ip, port, timeout) if banner else ""
            return port, True, b
        return port, False, ""
    except:
        return port, False, ""

def resolve_target(target):
    try:
        ip = socket.gethostbyname(target)
        return ip
    except socket.gaierror:
        print(f"[!] Could not resolve host: {target}")
        sys.exit(1)

def parse_ports(port_arg):
    ports = []
    for part in port_arg.split(","):
        part = part.strip()
        if "-" in part:
            start, end = part.split("-")
            ports.extend(range(int(start), int(end) + 1))
        else:
            ports.append(int(part))
    return ports

def main():
    parser = argparse.ArgumentParser(
        description="TCP Port Scanner - for authorized targets only"
    )
    parser.add_argument("target", help="Host or IP to scan")
    parser.add_argument(
        "-p", "--ports",
        help="Ports to scan: 80, 1-1024, 22,80,443 (default: common ports)",
        default=None
    )
    parser.add_argument(
        "-t", "--threads",
        help="Number of threads (default: 50)",
        type=int, default=50
    )
    parser.add_argument(
        "--timeout",
        help="Connection timeout in seconds (default: 1)",
        type=float, default=1
    )
    parser.add_argument(
        "-b", "--banner",
        help="Attempt banner grabbing",
        action="store_true"
    )
    args = parser.parse_args()

    ip = resolve_target(args.target)
    ports = parse_ports(args.ports) if args.ports else COMMON_PORTS

    print(f"\n{'='*50}")
    print(f"  Target   : {args.target} ({ip})")
    print(f"  Ports    : {len(ports)} to scan")
    print(f"  Threads  : {args.threads}")
    print(f"  Started  : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*50}\n")

    open_ports = []

    with ThreadPoolExecutor(max_workers=args.threads) as executor:
        futures = {
            executor.submit(scan_port, ip, port, args.timeout, args.banner): port
            for port in ports
        }
        for future in as_completed(futures):
            port, is_open, banner = future.result()
            if is_open:
                open_ports.append((port, banner))

    open_ports.sort()

    if open_ports:
        print(f"{'PORT':<10} {'STATE':<10} {'BANNER'}")
        print("-" * 50)
        for port, banner in open_ports:
            print(f"{port:<10} {'OPEN':<10} {banner}")
    else:
        print("[*] No open ports found.")

    print(f"\n[*] Scan complete. {len(open_ports)} open port(s) found.")
    print(f"[*] Finished : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

if __name__ == "__main__":
    main()
