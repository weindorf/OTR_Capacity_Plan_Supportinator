import sys
import os

# Add the project root to the Python path
project_root = os.path.abspath(os.path.dirname(__file__))
sys.path.insert(0, project_root)

from otr_supportinator.main import main

if __name__ == '__main__':
    main()