# -*- coding: utf-8 -*-

import sys
from pathlib import Path
import time

# Add lib folder to Python path
plugindir = Path.absolute(Path(__file__).parent)
paths = (".", "lib")
sys.path = [str(plugindir / p) for p in paths] + sys.path

import subprocess
import os
import logging
import json
from typing import List, Dict, Any
from flowlauncher import FlowLauncher

# Set up logging
log_file = os.path.join(plugindir, 'choco_plugin.log')
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

class ChocolateyPlugin(FlowLauncher):
    CACHE_FILE = plugindir / 'installed_cache.json'
    CACHE_TTL = 3600  # seconds (1 hour)

    def __init__(self):
        try:
            super().__init__()
        except Exception as e:
            logging.error(f"Error initializing plugin: {str(e)}")
            raise

    def invalidate_cache(self):
        try:
            if self.CACHE_FILE.exists():
                self.CACHE_FILE.unlink()
        except Exception as e:
            logging.error(f"Error invalidating cache: {str(e)}")

    def query(self, query: str) -> List[Dict[str, Any]]:
        try:
            if not query:
                # Show installed packages by default
                return self.list_installed_packages()

            # If there's a query, search for packages
            return self.search_packages(query)

        except Exception as e:
            logging.error(f"Error in query: {str(e)}")
            return [{
                "Title": "Error",
                "SubTitle": f"An error occurred: {str(e)}",
                "IcoPath": "Images/error.png"
            }]

    def _run_powershell_command(self, command: str) -> tuple[str, str, int]:
        """Run a PowerShell command and return stdout, stderr, and return code."""
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

            process = subprocess.Popen(
                ["powershell", "-Command", command],
                startupinfo=startupinfo,
                creationflags=subprocess.CREATE_NO_WINDOW,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            
            stdout, stderr = process.communicate()
            return stdout.strip(), stderr.strip(), process.returncode
        except Exception as e:
            logging.error(f"Error running PowerShell command: {str(e)}")
            return "", str(e), 1

    def install_package(self, package_name: str):
        try:
            self.invalidate_cache()
            # Run PowerShell as admin
            command = f"Start-Process powershell -Verb RunAs -ArgumentList '-Command \"choco install {package_name} -y\"'"
            subprocess.Popen(
                ["powershell", "-Command", command],
                creationflags=subprocess.CREATE_NEW_CONSOLE
            )
            self.show_msg(f"Starting installation of {package_name}")
            
            # Wait a bit for the installation to complete and update cache
            time.sleep(2)
            self.list_installed_packages()  # This will update the cache
                
        except Exception as e:
            logging.error(f"Error installing {package_name}: {str(e)}")
            self.show_msg(f"Error installing {package_name}: {str(e)}")

    def uninstall_package(self, package_name: str):
        try:
            self.invalidate_cache()
            # Run PowerShell as admin
            command = f"Start-Process powershell -Verb RunAs -ArgumentList '-Command \"choco uninstall {package_name} -y\"'"
            subprocess.Popen(
                ["powershell", "-Command", command],
                creationflags=subprocess.CREATE_NEW_CONSOLE
            )
            self.show_msg(f"Starting uninstallation of {package_name}")
            
            # Wait a bit for the uninstallation to complete and update cache
            time.sleep(2)
            self.list_installed_packages()  # This will update the cache
                
        except Exception as e:
            logging.error(f"Error uninstalling {package_name}: {str(e)}")
            self.show_msg(f"Error uninstalling {package_name}: {str(e)}")

    def upgrade_package(self, package_name: str):
        try:
            self.invalidate_cache()
            # Run PowerShell as admin
            command = f"Start-Process powershell -Verb RunAs -ArgumentList '-Command \"choco upgrade {package_name} -y\"'"
            subprocess.Popen(
                ["powershell", "-Command", command],
                creationflags=subprocess.CREATE_NEW_CONSOLE
            )
            self.show_msg(f"Starting upgrade of {package_name}")
                
        except Exception as e:
            logging.error(f"Error upgrading {package_name}: {str(e)}")
            self.show_msg(f"Error upgrading {package_name}: {str(e)}")

    def search_packages(self, query: str) -> List[Dict[str, Any]]:
        try:
            # First get basic search results
            command = f"& choco search {query} --limit-output"
            stdout, stderr, returncode = self._run_powershell_command(command)
            
            if returncode != 0:
                error_msg = stderr or stdout
                logging.error(f"Error searching packages: {error_msg}")
                return [{
                    "Title": "Error searching packages",
                    "SubTitle": error_msg,
                    "IcoPath": "Images/error.png"
                }]

            packages = []
            for line in stdout.splitlines():
                if line.strip():
                    try:
                        # Parse basic package info
                        parts = line.split('|')
                        name = parts[0]
                        version = parts[1] if len(parts) > 1 else "latest"
                        
                        # Get detailed info for each package
                        detail_command = f"& choco info {name} --limit-output"
                        detail_stdout, detail_stderr, detail_returncode = self._run_powershell_command(detail_command)
                        
                        description = ""
                        downloads = "0"
                        tags = ""
                        
                        if detail_returncode == 0 and detail_stdout:
                            detail_parts = detail_stdout.split('|')
                            if len(detail_parts) > 2:
                                description = detail_parts[2]
                            if len(detail_parts) > 3:
                                downloads = detail_parts[3]
                            if len(detail_parts) > 4:
                                tags = detail_parts[4]
                        
                        # Format subtitle with more details
                        subtitle_parts = []
                        if description:
                            subtitle_parts.append(description[:60] + "..." if len(description) > 60 else description)
                        if downloads and downloads != "0":
                            subtitle_parts.append(f"📥 {downloads} downloads")
                        if tags:
                            subtitle_parts.append(f"🏷️ {tags}")
                        
                        subtitle = " | ".join(subtitle_parts) if subtitle_parts else "No additional information available"
                        
                        packages.append({
                            "Title": f"{name} ({version})",
                            "SubTitle": subtitle,
                            "IcoPath": "Images/install.png",
                            "JsonRPCAction": {
                                "method": "install_package",
                                "parameters": [name],
                                "dontHideAfterAction": False
                            }
                        })
                    except Exception as e:
                        logging.error(f"Error parsing package info: {str(e)}")
                        continue

            if not packages:
                return [{
                    "Title": "No packages found",
                    "SubTitle": f"No results found for '{query}'",
                    "IcoPath": "Images/error.png"
                }]

            return packages
        except Exception as e:
            logging.error(f"Error searching packages: {str(e)}")
            return [{
                "Title": "Error searching packages",
                "SubTitle": str(e),
                "IcoPath": "Images/error.png"
            }]

    def list_installed_packages(self) -> List[Dict[str, Any]]:
        try:
            # Check cache
            if self.CACHE_FILE.exists():
                try:
                    with open(self.CACHE_FILE, 'r', encoding='utf-8') as f:
                        cache = json.load(f)
                    cache_time = cache.get('time', 0)
                    if time.time() - cache_time < self.CACHE_TTL:
                        return cache['packages']
                except Exception as e:
                    logging.error(f"Error reading cache: {str(e)}")
                    # If cache is bad, ignore and fetch fresh

            # Use Chocolatey's PowerShell API with better formatting
            command = "& choco list --limit-output --local-only"
            stdout, stderr, returncode = self._run_powershell_command(command)
            if returncode != 0:
                error_msg = stderr or stdout
                logging.error(f"Error listing packages: {error_msg}")
                return [{
                    "Title": "Error listing packages",
                    "SubTitle": error_msg,
                    "IcoPath": "Images/error.png"
                }]

            packages = []
            for line in stdout.splitlines():
                if line.strip():
                    try:
                        parts = line.split('|')
                        name = parts[0]
                        version = parts[1] if len(parts) > 1 else "unknown"
                        packages.append({
                            "Title": f"{name} ({version})",
                            "SubTitle": "Press Enter to uninstall",
                            "IcoPath": "Images/install.png",
                            "JsonRPCAction": {
                                "method": "uninstall_package",
                                "parameters": [name],
                                "dontHideAfterAction": False
                            }
                        })
                    except Exception as e:
                        logging.error(f"Error parsing package info: {str(e)}")
                        continue

            if not packages:
                return [{
                    "Title": "No packages installed",
                    "SubTitle": "Type a package name to search and install",
                    "IcoPath": "Images/choco.png"
                }]

            # Save to cache
            try:
                with open(self.CACHE_FILE, 'w', encoding='utf-8') as f:
                    json.dump({'time': time.time(), 'packages': packages}, f)
            except Exception as e:
                logging.error(f"Error writing cache: {str(e)}")

            return packages
        except Exception as e:
            logging.error(f"Error listing packages: {str(e)}")
            return [{
                "Title": "Error listing packages",
                "SubTitle": str(e),
                "IcoPath": "Images/error.png"
            }]

    def show_msg(self, msg: str):
        self.show_msg(msg)

if __name__ == "__main__":
    try:
        ChocolateyPlugin()
    except Exception as e:
        logging.critical(f"Critical error in plugin: {str(e)}")
        print(json.dumps({
            "result": [{
                "Title": "Critical Error",
                "SubTitle": "The plugin encountered a critical error. Check the log file for details.",
                "IcoPath": "Images/error.png"
            }]
        })) 