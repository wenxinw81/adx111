from adx_report_agent.web_app import analysis_name


def test_analysis_name() -> None:
    assert analysis_name("basic") == "基础分析"
    assert analysis_name("spend") == "花销专门分析"
    assert analysis_name("bidding") == "竞价专门分析"
