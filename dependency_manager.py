#!/usr/bin/env python3
"""
Dependency Manager for Python Projects
---------------------------------------
This script helps manage dependencies and imports in Python projects by:
1. Checking for outdated packages
2. Finding unused imports in Python files
3. Suggesting updates to requirements.txt
"""

import os
import sys
import importlib
import subprocess
import re
import pkg_resources
from pathlib import Path

def check_outdated_packages():
    """Check for outdated packages in the current environment."""
    print("Checking for outdated packages...")
    subprocess.run([sys.executable, "-m", "pip", "list", "--outdated"])
    
def update_packages():
    """Update all packages to their latest versions."""
    answer = input("Do you want to update all packages to their latest versions? (y/n): ")
    if answer.lower() == 'y':
        subprocess.run([sys.executable, "-m", "pip", "install", "--upgrade", "-r", "requirements.txt"])
        print("All packages have been updated!")
    else:
        print("Update canceled.")

def find_imports_in_file(file_path):
    """Find all imports in a Python file."""
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Match import statements
    import_pattern = r'^import\s+([a-zA-Z0-9_,\s]+)$|^from\s+([a-zA-Z0-9_.]+)\s+import\s+([a-zA-Z0-9_,\s*]+)$'
    imports = []
    
    # Also look for fallback import patterns (try/except blocks)
    fallback_pattern = r'try:\s*import\s+([a-zA-Z0-9_,\s]+)|try:\s*from\s+([a-zA-Z0-9_.]+)\s+import'
    fallback_imports = []
    
    for line in content.split('\n'):
        line = line.strip()
        match = re.match(import_pattern, line)
        if match:
            if match.group(1):  # import x
                modules = [m.strip() for m in match.group(1).split(',')]
                imports.extend(modules)
            elif match.group(2) and match.group(3):  # from x import y
                module = match.group(2)
                if match.group(3) == '*':
                    imports.append(module)
                else:
                    imports.append(module)
    
    # Add fallback imports detection
    fallback_matches = re.findall(fallback_pattern, content, re.MULTILINE)
    for match in fallback_matches:
        if match[0]:  # try: import x
            modules = [m.strip() for m in match[0].split(',')]
            fallback_imports.extend(modules)
        elif match[1]:  # try: from x import
            fallback_imports.append(match[1])
    
    return imports, fallback_imports

def scan_project_imports(directory='.'):
    """Scan all Python files in a project for imports."""
    all_imports = set()
    all_fallback_imports = set()
    
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith('.py'):
                file_path = os.path.join(root, file)
                try:
                    imports, fallback_imports = find_imports_in_file(file_path)
                    all_imports.update(imports)
                    all_fallback_imports.update(fallback_imports)
                except Exception as e:
                    print(f"Error processing {file_path}: {e}")
    
    return all_imports, all_fallback_imports

def get_installed_packages():
    """Get all installed packages in the current environment."""
    return {pkg.key for pkg in pkg_resources.working_set}

def parse_requirements_file(requirements_path):
    """Parse a requirements.txt file accounting for commented optional dependencies."""
    if not os.path.exists(requirements_path):
        print(f"Requirements file {requirements_path} not found!")
        return {}, {}
    
    with open(requirements_path, 'r', encoding='utf-8') as f:
        requirements = f.readlines()
    
    required_packages = {}
    optional_packages = {}
    
    for req in requirements:
        req = req.strip()
        if not req:
            continue
            
        if req.startswith('# ') and '>=' in req:
            # This is a commented optional dependency
            # Extract package info from comment
            package_line = req[2:].strip()  # Remove '# ' prefix
            package_name = package_line.split('>=')[0].split('==')[0].split('>')[0].split('<')[0].strip()
            optional_packages[package_name] = package_line
        elif not req.startswith('#'):
            # Regular required dependency
            package_name = req.split('>=')[0].split('==')[0].split('>')[0].split('<')[0].strip()
            required_packages[package_name] = req
    
    return required_packages, optional_packages

def find_unused_dependencies(imports, fallback_imports, requirements_path='requirements.txt'):
    """Find packages in requirements.txt that are not imported in the code."""
    required_packages, optional_packages = parse_requirements_file(requirements_path)
    
    # Convert import names to probable package names (simplistic approach)
    potential_packages = set()
    for imp in imports:
        # Take the first part of the import, which often corresponds to package name
        potential_packages.add(imp.split('.')[0].lower())
    
    # Handle fallback packages differently
    for imp in fallback_imports:
        # These are imports wrapped in try/except, so they are optional
        potential_packages.add(imp.split('.')[0].lower())
    
    # Find packages in requirements that aren't imported
    unused_required = {pkg: ver for pkg, ver in required_packages.items() 
                      if pkg.lower() not in potential_packages}
    
    unused_optional = {pkg: ver for pkg, ver in optional_packages.items() 
                      if pkg.lower() not in potential_packages}
    
    return unused_required, unused_optional

def find_missing_dependencies(imports, fallback_imports, requirements_path='requirements.txt'):
    """Find imports that might be missing from requirements.txt."""
    installed_packages = {pkg.lower() for pkg in get_installed_packages()}
    required_packages, optional_packages = parse_requirements_file(requirements_path)
    
    # Combined required and optional for checking what's already in requirements
    all_required = {pkg.lower() for pkg in required_packages.keys()}
    all_required.update({pkg.lower() for pkg in optional_packages.keys()})
    
    # Convert import names to probable package names 
    missing_required = set()
    could_be_optional = set()
    
    # Process main imports as potentially required
    for imp in imports:
        package_name = imp.split('.')[0].lower()
        # Check if it's an installed package (not standard library)
        if package_name in installed_packages and package_name not in all_required:
            missing_required.add(package_name)
    
    # Process fallback imports as potentially optional
    for imp in fallback_imports:
        package_name = imp.split('.')[0].lower()
        # Check if it's an installed package (not standard library)
        if package_name in installed_packages and package_name not in all_required:
            could_be_optional.add(package_name)
    
    return missing_required, could_be_optional

def check_package_exists(package_name):
    """Check if a package exists on PyPI."""
    try:
        subprocess.check_output([sys.executable, "-m", "pip", "install", "--quiet", "--dry-run", package_name])
        return True
    except subprocess.CalledProcessError:
        return False

def update_requirements_file(requirements_path, packages_to_add=None, packages_to_comment=None):
    """Update the requirements.txt file with new packages and comment out unused ones."""
    if not os.path.exists(requirements_path):
        print(f"Requirements file {requirements_path} not found!")
        return False
    
    with open(requirements_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    # Convert packages to comment to lowercase for case-insensitive matching
    if packages_to_comment:
        packages_to_comment = {pkg.lower(): comment for pkg, comment in packages_to_comment.items()}
    
    # Process existing lines
    new_lines = []
    for line in lines:
        stripped = line.strip()
        if not stripped or stripped.startswith('#'):
            new_lines.append(line)  # Keep comments and blank lines as is
            continue
            
        # Handle actual package specifications
        package_name = stripped.split('>=')[0].split('==')[0].split('>')[0].split('<')[0].strip().lower()
        
        if packages_to_comment and package_name in packages_to_comment:
            # Comment out this package
            new_lines.append(f"# {stripped}  # {packages_to_comment[package_name]}\n")
        else:
            new_lines.append(line)
    
    # Add new packages
    if packages_to_add:
        new_lines.append("\n# Newly detected dependencies\n")
        for package_name in sorted(packages_to_add):
            # Try to get the installed version
            try:
                pkg_info = next(p for p in pkg_resources.working_set if p.key.lower() == package_name.lower())
                new_lines.append(f"{pkg_info.key}>={pkg_info.version}\n")
            except StopIteration:
                # If we can't find the version, just add the name
                new_lines.append(f"{package_name}\n")
    
    # Write the updated file
    with open(requirements_path, 'w', encoding='utf-8') as f:
        f.writelines(new_lines)
    
    return True

def main():
    print("Python Dependency Manager")
    print("-" * 50)
    
    # Check for outdated packages
    check_outdated_packages()
    print("-" * 50)
    
    # Scan project for imports
    project_dir = input("Enter project directory path (default: current directory): ") or '.'
    requirements_path = input("Enter requirements.txt path (default: ./requirements.txt): ") or 'requirements.txt'
    
    print(f"Scanning Python files in {project_dir}...")
    imports, fallback_imports = scan_project_imports(project_dir)
    print(f"Found {len(imports)} regular imports and {len(fallback_imports)} fallback/optional imports.")
    
    # Find unused dependencies
    unused_required, unused_optional = find_unused_dependencies(imports, fallback_imports, requirements_path)
    
    if unused_required:
        print("\nPotentially unused required packages:")
        for pkg, ver in unused_required.items():
            print(f"  - {ver}")
    else:
        print("\nNo potentially unused required packages found!")
    
    if unused_optional:
        print("\nPotentially unused optional packages:")
        for pkg, ver in unused_optional.items():
            print(f"  - {ver} (already marked as optional)")
    
    # Find missing dependencies
    missing_required, could_be_optional = find_missing_dependencies(imports, fallback_imports, requirements_path)
    
    if missing_required:
        print("\nPotentially missing required packages:")
        for pkg in missing_required:
            if check_package_exists(pkg):
                print(f"  - {pkg}")
            else:
                print(f"  - {pkg} (might be a local package)")
    else:
        print("\nNo potentially missing required packages detected!")
    
    if could_be_optional:
        print("\nPotentially missing optional packages:")
        for pkg in could_be_optional:
            if check_package_exists(pkg):
                print(f"  - {pkg}")
            else:
                print(f"  - {pkg} (might be a local package)")
    
    # Ask if they want to update the requirements file
    if missing_required or unused_required:
        answer = input("\nDo you want to update requirements.txt with these findings? (y/n): ")
        if answer.lower() == 'y':
            # Add missing required packages
            packages_to_add = missing_required if missing_required else None
            
            # Comment out unused required packages
            packages_to_comment = {}
            if unused_required:
                for pkg in unused_required:
                    packages_to_comment[pkg] = "Unused dependency"
            
            if update_requirements_file(requirements_path, packages_to_add, packages_to_comment):
                print(f"Requirements file {requirements_path} has been updated!")
            else:
                print("Failed to update requirements file.")
    
    # Offer to update all packages
    if os.path.exists(requirements_path):
        print("-" * 50)
        update_packages()

if __name__ == "__main__":
    main()