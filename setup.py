# Create a Python package for our modules
import os

# Create package directory if it doesn't exist
os.makedirs('chaiotic', exist_ok=True)

# Create __init__.py to make the directory a proper package
with open('chaiotic/__init__.py', 'w') as f:
    f.write('"""Grammar and logic checking toolkit for documents."""\n\n')
    f.write('__version__ = "0.1.0"\n')

print("Created chaiotic package structure")