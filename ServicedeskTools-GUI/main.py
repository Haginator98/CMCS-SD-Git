#!/usr/bin/env python3
"""Servicedesk Tools — GUI launcher for Azure/Exchange PowerShell scripts."""

from app import ServicedeskApp


def main():
    app = ServicedeskApp()
    app.mainloop()


if __name__ == "__main__":
    main()
