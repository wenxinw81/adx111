import os

os.environ.setdefault("ADX_AGENT_CONFIG", "configs/agent.runtime.example.json")
os.environ.setdefault("ADX_SSH_ENABLED", "true")

from adx_report_agent.desktop_launcher import main


if __name__ == "__main__":
    main()
