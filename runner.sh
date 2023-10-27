#!/bin/bash

# Get the absolute path of the script
SCRIPT_PATH="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

# Check if pything is already installed
if command -v python3 &>/dev/null; then
  echo "Python already installed"
else
  echo "Installing python"
  brew install python
fi

if [["$PATH" != *"/usr/local/bin"*]]; then
  echo "Adding Python directory to PATH"
  echo 'export PATH="$PATH:/usr/local/bin"' >> "$SCRIPT_PATH/.bash_profile"
  source "$SCRIPT_PATH/.bash_profile"
fi

python_version=$(python3 --version 2>&1)
echo "Python version: $python_version"

if command -v python3 &>/dev/null; then
  GUI_SCRIPT="$SCRIPT_PATH/orderExtractor/"

  if [ -d "$GUI_SCRIPT" ]; then

    echo changing directory to project
    cd "$GUI_SCRIPT"

    sudo apt update -y 

    sudo apt install python3-tk -y

    echo installing dependencies
    python3 -m pip install -r requirements.txt

    echo "Executing gui.py"
    python3 gui.py

  else
    echo "gui.py not found"   
  fi
else
  echo "Python is not installed correctly. Can not run program"
fi

