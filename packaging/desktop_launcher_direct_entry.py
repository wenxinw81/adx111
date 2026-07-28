import os

os.environ.setdefault("ADX_AGENT_CONFIG", "configs/agent.direct.example.json")

from adx_report_agent.desktop_launcher import main


if __name__ == "__main__":
    main()
