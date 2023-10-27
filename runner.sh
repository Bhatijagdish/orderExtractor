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


if [[":$PATH:" != *":/usr/local/bin/:"*]]; then
  echo "Adding Python directory to PATH"
  echo 'export PATH="$PATH:/usr/local/bin"' >> "$SCRIPT_DIR/.bash_profile"
  source "$SCRIPT_DIR/.bash_profile"
fi

python_version=$(python3 --version 2>&1)
echo "Python version: $python_version"

if command -v python3 &>/dev/null; then
  GUI_SCRIPT="$SCRIPT_DIR/orderExtractor/gui.py"
  if [ -f "$GUI_SCRIPT" ]; then
    echo "Executing gui.py"
    python3 "$GUI_SCRIPT"
  else
    echo "gui.py not found"
  fi
else
  echo "Python is not installed correctly. Can not run program"
fi

