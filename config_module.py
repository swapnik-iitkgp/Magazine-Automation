#!/usr/bin/env python
# config_module.py

import json
import os

def load_config(config_path):
    """Load configuration settings from a JSON file."""
    if not os.path.exists(config_path):
        return {}
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)

def save_config(config, config_path):
    """Save the configuration dictionary back to the JSON file."""
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4)
    print(f"[INFO] Configuration saved to: {config_path}")

if __name__ == "__main__":
    # Example usage
    cfg_path = "config.json"
    cfg = load_config(cfg_path)
    # For demonstration, ensure we have a template_file key
    cfg.setdefault("template_file", "template.indd")
    save_config(cfg, cfg_path)
