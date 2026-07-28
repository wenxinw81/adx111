from __future__ import annotations

import os
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / ".vendor_pymysql"))
import pymysql  # noqa: E402


DAY = os.getenv("REPORT_DAY", "2026-07-25")
PREV_DAY = os.getenv(
    "REPORT_PREV_DAY",
    (datetime.strptime(DAY, "%Y-%m-%d") - timedelta(days=1)).strftime("%Y-%m-%d"),
)
_day_dt = datetime.strptime(DAY, "%Y-%m-%d")
OUT = Path(
    os.getenv(
        "REPORT_OUTPUT",
        f"outputs/adx_basic_report_{_day_dt:%Y%m%d}/{_day_dt.month}.{_day_dt.day}广告投放基础分析.xlsx",
    )
)
DEFAULT_APP_MAPPING_CSV = Path(__file__).resolve().parents[1] / "data" / "app_mapping.csv"
APP_MAPPING_CSV = Path(os.getenv("APP_MAPPING_CSV", str(DEFAULT_APP_MAPPING_CSV)))
ORDER_ID_TEXT = os.getenv("REPORT_ORDER_ID", "").strip()
ORDER_ID = int(ORDER_ID_TEXT) if ORDER_ID_TEXT else None


def order_filter(alias: str = "r") -> str:
    return f" and {alias}.order_id = {ORDER_ID}" if ORDER_ID is not None else ""


def request_order_exists() -> str:
    if ORDER_ID is None:
        return ""
    return f"""
      and exists (
        select 1
        from ods.ods_ad_adx_bid_response_rt r
        where r.response_id=request_id
          and r.bid_imp_id=imp_id
          and r.order_id={ORDER_ID}
          and r.response_time >= request_time
          and r.response_time < date_add(request_time, interval 1 day)
      )
    """


def conn() -> pymysql.Connection:
    password = os.getenv("DORIS_PASSWORD")
    if not password:
        raise RuntimeError("DORIS_PASSWORD is required")
    return pymysql.connect(
        host=os.getenv("DORIS_HOST", "127.0.0.1"),
        port=int(os.getenv("DORIS_PORT", "19030")),
        user=os.getenv("DORIS_USER", "WishFox"),
        password=password,
        database=os.getenv("DORIS_DATABASE", "ads"),
        connect_timeout=10,
        read_timeout=int(os.getenv("DORIS_READ_TIMEOUT", "900")),
        write_timeout=int(os.getenv("DORIS_WRITE_TIMEOUT", "900")),
        charset="utf8mb4",
    )


def q(sql: str, params: tuple = ()) -> pd.DataFrame:
    attempts = int(os.getenv("DORIS_QUERY_RETRIES", "3"))
    for attempt in range(1, attempts + 1):
        try:
            with conn() as c:
                return pd.read_sql(sql, c, params=params)
        except pymysql.err.OperationalError as exc:
            code = exc.args[0] if exc.args else None
            if code not in {2006, 2013} or attempt >= attempts:
                raise
            wait_seconds = min(2**attempt, 8)
            print(f"Doris query lost connection ({code}); retry {attempt}/{attempts - 1} after {wait_seconds}s", file=sys.stderr)
            time.sleep(wait_seconds)
    raise RuntimeError("Doris query failed without returning a result")


def pct(num: pd.Series, den: pd.Series) -> pd.Series:
    return (num / den.replace(0, pd.NA)).fillna(0)


def add_rates(df: pd.DataFrame, has_invite: bool = True) -> pd.DataFrame:
    df = df.copy()
    if has_invite and "邀约" in df:
        df["邀约→出价"] = pct(df["出价"], df["邀约"])
    df["出价→竞得"] = pct(df["竞得"], df["出价"])
    df["竞得→点击"] = pct(df["点击"], df["竞得"])
    df["点击→落地页曝光"] = pct(df["落地页曝光"], df["点击"])
    df["落地页曝光→跳转"] = pct(df["跳转"], df["落地页曝光"])
    return df


def spend_expr() -> str:
    return "sum(coalesce(cast(nullif(wpr, '') as double), 0)) / 100000"


def load_app_mapping() -> dict[str, str]:
    if not APP_MAPPING_CSV.exists():
        return {}
    df = pd.read_csv(APP_MAPPING_CSV, encoding="utf-8-sig")
    required = {"app_package_name", "app_name"}
    if not required.issubset(df.columns):
        return {}
    df = df[["app_package_name", "app_name", "is_verified", "request_count", "etl_time"]].copy()
    df["app_package_name"] = df["app_package_name"].fillna("").astype(str).str.strip()
    df["app_name"] = df["app_name"].fillna("").astype(str).str.strip()
    df = df[(df["app_package_name"].ne("")) & (df["app_name"].ne(""))]
    df["is_verified"] = pd.to_numeric(df.get("is_verified"), errors="coerce").fillna(0)
    df["request_count"] = pd.to_numeric(df.get("request_count"), errors="coerce").fillna(0)
    df = df.sort_values(["is_verified", "request_count", "etl_time"], ascending=[False, False, False])
    return df.drop_duplicates("app_package_name").set_index("app_package_name")["app_name"].to_dict()


def main() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)

    daily_sql = f"""
    with
    req as (
      select date(request_time) d, count(*) invite
      from ods.ods_ad_adx_bid_request_rt
      where request_time >= %s and request_time < date_add(%s, interval 2 day)
      group by date(request_time)
    ),
    bid as (
      select date(response_time) d, count(*) bid
      from ods.ods_ad_adx_bid_response_rt
      where response_time >= %s and response_time < date_add(%s, interval 2 day)
      group by date(response_time)
    ),
    win as (
      select date(response_time) d, count(*) win, {spend_expr()} spend
      from ods.ods_ad_adx_bid_response_result_rt
      where response_time >= %s and response_time < date_add(%s, interval 2 day)
      group by date(response_time)
    ),
    clk as (
      select date(response_time) d, count(*) click
      from ods.ods_ad_adx_bid_response_click_rt
      where response_time >= %s and response_time < date_add(%s, interval 2 day)
      group by date(response_time)
    ),
    clk_pairs as (
      select date(response_time) d, response_id, bid_seat_id
      from ods.ods_ad_adx_bid_response_click_rt
      where response_time >= %s and response_time < date_add(%s, interval 2 day)
      group by date(response_time), response_id, bid_seat_id
    ),
    track_keys as (
      select
        coalesce(nullif(regexp_extract(cast(t.preset_attributes as string), '[?&]rid=([^&]+)', 1), ''), t.root_id) response_id_key,
        coalesce(nullif(regexp_extract(cast(t.preset_attributes as string), '[?&]pyck=([^&]+)', 1), ''), t.parent_event_id) bid_seat_id_key,
        t.event_code
      from ods.ods_track_event_default_raw t
      where t.event_time >= %s and t.event_time < date_add(%s, interval 2 day)
    ),
    trk as (
      select c.d,
        count(distinct case when t.event_code='pageExp' then concat(t.response_id_key,'#',t.bid_seat_id_key) end) page_exp,
        count(distinct case when t.event_code='jumpButton' then concat(t.response_id_key,'#',t.bid_seat_id_key) end) jump,
        count(distinct case when t.event_code in ('pageLoadFail','pageLoadTimeout') then concat(t.response_id_key,'#',t.bid_seat_id_key) end) load_fail
      from clk_pairs c
      join track_keys t
        on t.response_id_key=c.response_id and t.bid_seat_id_key=c.bid_seat_id
      group by c.d
    )
    select req.d 日期, invite 邀约, coalesce(bid,0) 出价, coalesce(win,0) 竞得,
           coalesce(click,0) 点击, coalesce(page_exp,0) 落地页曝光,
           coalesce(jump,0) 跳转, coalesce(spend,0) `花销(元)`, coalesce(load_fail,0) 页面加载失败
    from req
    left join bid on req.d=bid.d
    left join win on req.d=win.d
    left join clk on req.d=clk.d
    left join trk on req.d=trk.d
    order by req.d
    """
    daily_params = (
        PREV_DAY,
        PREV_DAY,
        PREV_DAY,
        PREV_DAY,
        PREV_DAY,
        PREV_DAY,
        PREV_DAY,
        PREV_DAY,
        PREV_DAY,
        PREV_DAY,
        PREV_DAY,
        PREV_DAY,
    )
    if ORDER_ID is not None:
        daily_sql = f"""
        with
        base as (
          select date(r.response_time) d, r.response_id, r.bid_seat_id
          from ods.ods_ad_adx_bid_response_rt r
          where r.response_time >= %s and r.response_time < date_add(%s, interval 2 day)
            and r.order_id = {ORDER_ID}
        ),
        req as (
          select d, count(*) invite
          from base
          group by d
        ),
        bid as (
          select d, count(*) bid
          from base
          group by d
        ),
        win as (
          select b.d, count(*) win, sum(coalesce(cast(nullif(w.wpr, '') as double), 0)) / 100000 spend
          from base b
          join ods.ods_ad_adx_bid_response_result_rt w
            on b.response_id=w.response_id and b.bid_seat_id=w.bid_seat_id
          where w.response_time >= %s and w.response_time < date_add(%s, interval 2 day)
          group by b.d
        ),
        clk_pairs as (
          select b.d, c.response_id, c.bid_seat_id
          from base b
          join ods.ods_ad_adx_bid_response_click_rt c
            on b.response_id=c.response_id and b.bid_seat_id=c.bid_seat_id
          where c.response_time >= %s and c.response_time < date_add(%s, interval 2 day)
          group by b.d, c.response_id, c.bid_seat_id
        ),
        clk as (
          select d, count(*) click
          from clk_pairs
          group by d
        ),
        track_keys as (
          select
            coalesce(nullif(regexp_extract(cast(t.preset_attributes as string), '[?&]rid=([^&]+)', 1), ''), t.root_id) response_id_key,
            coalesce(nullif(regexp_extract(cast(t.preset_attributes as string), '[?&]pyck=([^&]+)', 1), ''), t.parent_event_id) bid_seat_id_key,
            t.event_code
          from ods.ods_track_event_default_raw t
          where t.event_time >= %s and t.event_time < date_add(%s, interval 2 day)
        ),
        trk as (
          select c.d,
            count(distinct case when t.event_code='pageExp' then concat(t.response_id_key,'#',t.bid_seat_id_key) end) page_exp,
            count(distinct case when t.event_code='jumpButton' then concat(t.response_id_key,'#',t.bid_seat_id_key) end) jump,
            count(distinct case when t.event_code in ('pageLoadFail','pageLoadTimeout') then concat(t.response_id_key,'#',t.bid_seat_id_key) end) load_fail
          from clk_pairs c
          join track_keys t
            on t.response_id_key=c.response_id and t.bid_seat_id_key=c.bid_seat_id
          group by c.d
        )
        select req.d 日期, invite 邀约, coalesce(bid,0) 出价, coalesce(win,0) 竞得,
               coalesce(click,0) 点击, coalesce(page_exp,0) 落地页曝光,
               coalesce(jump,0) 跳转, coalesce(spend,0) `花销(元)`, coalesce(load_fail,0) 页面加载失败
        from req
        left join bid on req.d=bid.d
        left join win on req.d=win.d
        left join clk on req.d=clk.d
        left join trk on req.d=trk.d
        order by req.d
        """
        daily_params = (PREV_DAY, PREV_DAY, PREV_DAY, PREV_DAY, PREV_DAY, PREV_DAY, PREV_DAY, PREV_DAY)
    daily = q(
        daily_sql,
        daily_params,
    )
    daily = add_rates(daily)
    current = daily[daily["日期"].astype(str) == DAY].copy()
    prev = daily[daily["日期"].astype(str) == PREV_DAY].copy()

    wide_sql = f"""
    with
    win as (
      select response_id, bid_seat_id, count(*) win, sum(coalesce(cast(nullif(wpr, '') as double),0))/100000 spend_yuan
      from ods.ods_ad_adx_bid_response_result_rt
      where response_time >= %s and response_time < date_add(%s, interval 1 day)
      group by response_id, bid_seat_id
    ),
    clk as (
      select response_id, bid_seat_id, count(*) click
      from ods.ods_ad_adx_bid_response_click_rt
      where response_time >= %s and response_time < date_add(%s, interval 1 day)
      group by response_id, bid_seat_id
    ),
    track_keys as (
      select
        coalesce(nullif(regexp_extract(cast(t.preset_attributes as string), '[?&]rid=([^&]+)', 1), ''), t.root_id) response_id_key,
        coalesce(nullif(regexp_extract(cast(t.preset_attributes as string), '[?&]pyck=([^&]+)', 1), ''), t.parent_event_id) bid_seat_id_key,
        regexp_extract(cast(t.preset_attributes as string), '/dragIndex/([0-9]+)', 1) page_id,
        t.event_code
      from ods.ods_track_event_default_raw t
      where t.event_time >= %s and t.event_time < date_add(%s, interval 1 day)
    ),
    trk as (
      select t.response_id_key, t.bid_seat_id_key,
        max(t.page_id) page_id,
        max(case when t.event_code='pageExp' then 1 else 0 end) page_exp,
        max(case when t.event_code='pageViewComplete' then 1 else 0 end) complete,
        max(case when t.event_code in ('scrollStart','scrollEnd') then 1 else 0 end) scroll,
        max(case when t.event_code in ('vidPlay','vidPause') then 1 else 0 end) video,
        max(case when t.event_code='jumpButton' then 1 else 0 end) jump
      from track_keys t
      join clk c on t.response_id_key=c.response_id and t.bid_seat_id_key=c.bid_seat_id
      group by t.response_id_key, t.bid_seat_id_key
    ),
    dim_strategy as (
      select strategy_id, strategy_name, order_id, order_name
      from (
        select
          strategy_id,
          strategy_name,
          order_id,
          order_name,
          row_number() over(partition by strategy_id order by etl_time desc) rn
        from dim.dim_ad_dsp_strategy
      ) s
      where rn=1
    )
    select
      r.response_time,
      q.request_time,
      hour(q.request_time) hour,
      coalesce(nullif(q.app_bundle,''),'(空)') app_bundle,
      q.imp_bidfloor,
      r.response_id,
      r.bid_seat_id,
      r.bid_price,
      r.order_id,
      r.strategy_id,
      r.creative_id,
      coalesce(s.order_name, d.order_name, cast(r.order_id as string)) order_name,
      coalesce(s.strategy_name, d.strategy_name, cast(r.strategy_id as string)) strategy_name,
      coalesce(d.creative_name, cast(r.creative_id as string)) creative_name,
      d.material_id,
      d.material_type,
      r.land_page,
      r.curl,
      r.dpl_url,
      coalesce(win.win,0) win,
      coalesce(win.spend_yuan,0) spend_yuan,
      coalesce(clk.click,0) click,
      coalesce(trk.page_id,'') page_id,
      coalesce(trk.page_exp,0) page_exp,
      coalesce(trk.complete,0) complete,
      coalesce(trk.scroll,0) scroll,
      coalesce(trk.video,0) video,
      coalesce(trk.jump,0) jump
    from ods.ods_ad_adx_bid_response_rt r
    left join ods.ods_ad_adx_bid_request_rt q
      on q.request_id=r.response_id and q.imp_id=r.bid_imp_id
      and q.request_time >= %s and q.request_time < date_add(%s, interval 1 day)
    left join win on r.response_id=win.response_id and r.bid_seat_id=win.bid_seat_id
    left join clk on r.response_id=clk.response_id and r.bid_seat_id=clk.bid_seat_id
    left join trk on r.response_id=trk.response_id_key and r.bid_seat_id=trk.bid_seat_id_key
    left join dim_strategy s on r.strategy_id=s.strategy_id
    left join dim.dim_ad_dsp_creative d on trim(r.creative_id)=cast(d.creative_id as string)
    where r.response_time >= %s and r.response_time < date_add(%s, interval 1 day)
    {order_filter("r")}
    """
    wide = q(wide_sql, (DAY, DAY, DAY, DAY, DAY, DAY, DAY, DAY, DAY, DAY))
    app_map = load_app_mapping()
    wide["APP名称"] = wide["app_bundle"].map(app_map).fillna(wide["app_bundle"])
    wide["订单策略素材"] = (
        wide["order_name"].fillna("").astype(str)
        + " / "
        + wide["strategy_name"].fillna("").astype(str)
        + " / "
        + wide["creative_name"].fillna("").astype(str)
    )
    wide["直接跳转/无落地页"] = (
        wide["land_page"].fillna("").eq("")
        & wide["curl"].fillna("").ne("")
        & wide["page_exp"].eq(0)
    )

    hourly_req = q(
        f"""
        select hour(request_time) hour, count(*) 邀约, avg(imp_bidfloor) `平均底价`
        from ods.ods_ad_adx_bid_request_rt
        where request_time >= %s and request_time < date_add(%s, interval 1 day)
          {request_order_exists()}
        group by hour(request_time)
        order by hour
        """,
        (DAY, DAY),
    )

    top_apps = wide.groupby("app_bundle").size().sort_values(ascending=False).head(30).index.tolist()
    if "(空)" not in top_apps:
        top_apps.append("(空)")
    app_req_sql = f"""
        select coalesce(nullif(app_bundle,''),'(空)') app_bundle, count(*) 邀约, count(distinct hour(request_time)) 活跃小时
        from ods.ods_ad_adx_bid_request_rt
        where request_time >= %s and request_time < date_add(%s, interval 1 day)
          and coalesce(nullif(app_bundle,''),'(空)') in ({",".join(["%s"] * len(top_apps))})
          {request_order_exists()}
        group by coalesce(nullif(app_bundle,''),'(空)')
    """
    app_req = q(app_req_sql, (DAY, DAY, *top_apps))
    app_req["APP名称"] = app_req["app_bundle"].map(app_map).fillna(app_req["app_bundle"])

    def bid_agg(keys: list[str]) -> pd.DataFrame:
        g = wide.groupby(keys, dropna=False).agg(
            出价=("response_id", "size"),
            竞得=("win", "sum"),
            点击=("click", "sum"),
            落地页曝光=("page_exp", "sum"),
            跳转=("jump", "sum"),
            花销=("spend_yuan", "sum"),
            直接跳转订单出价=("直接跳转/无落地页", "sum"),
        ).reset_index()
        g = g.rename(columns={"spend_yuan": "花销(元)", "花销": "花销(元)"})
        return add_rates(g, has_invite=False)

    hourly_bid = bid_agg(["hour"])
    hourly = pd.merge(hourly_req, hourly_bid, on="hour", how="left").fillna(0)
    hourly = add_rates(hourly, has_invite=True)

    app_bid = bid_agg(["APP名称", "app_bundle"])
    app = pd.merge(app_req, app_bid, on=["APP名称", "app_bundle"], how="right").fillna(0)
    leading = ["APP名称", "app_bundle"]
    app = app[leading + [c for c in app.columns if c not in leading]]
    app = app.sort_values("出价", ascending=False).head(31)
    app = add_rates(app, has_invite=True)

    gantt = wide.pivot_table(index=["order_id", "strategy_id", "order_name", "strategy_name"], columns="hour", values="response_id", aggfunc="count", fill_value=0).reset_index()
    hour_cols = [c for c in range(24) if c in gantt.columns]
    summary = bid_agg(["order_id", "strategy_id"]).rename(columns={"花销(元)": "花销(元)"})
    first_last = wide.groupby(["order_id", "strategy_id"]).agg(开始时间=("response_time", "min"), 最后出价时间=("response_time", "max"), 活跃小时=("hour", "nunique")).reset_index()
    gantt = gantt.merge(first_last, on=["order_id", "strategy_id"], how="left").merge(summary[["order_id", "strategy_id", "出价", "竞得", "点击", "花销(元)"]], on=["order_id", "strategy_id"], how="left")
    gantt = gantt[["order_id", "strategy_id", "order_name", "strategy_name", "开始时间", "最后出价时间", "活跃小时", "出价", "竞得", "点击", "花销(元)"] + hour_cols].sort_values("出价", ascending=False)

    order_fit = bid_agg(["order_id", "order_name"]).sort_values("出价", ascending=False)
    strategy_fit = bid_agg(["order_id", "strategy_id", "order_name", "strategy_name"]).sort_values("出价", ascending=False)
    creative_fit = bid_agg(["order_id", "strategy_id", "creative_id", "order_name", "strategy_name", "creative_name", "material_id", "material_type"]).sort_values("出价", ascending=False)

    landing = wide.groupby("page_id", dropna=False).agg(
        访问数=("page_exp", "sum"),
        完读访问=("complete", "sum"),
        滚动访问=("scroll", "sum"),
        视频访问=("video", "sum"),
        页面点击访问=("jump", "sum"),
        点击后无落地页=("直接跳转/无落地页", "sum"),
    ).reset_index().rename(columns={"page_id": "落地页ID"})
    landing = landing[(landing["落地页ID"].ne("")) | (landing["访问数"].gt(0)) | (landing["点击后无落地页"].gt(0))]
    landing["完读率"] = pct(landing["完读访问"], landing["访问数"])
    landing["滚动率"] = pct(landing["滚动访问"], landing["访问数"])
    landing["视频率"] = pct(landing["视频访问"], landing["访问数"])
    landing["页面点击率"] = pct(landing["页面点击访问"], landing["访问数"])
    landing = landing.sort_values("访问数", ascending=False)

    app_hour = bid_agg(["APP名称", "app_bundle", "hour"]).sort_values(["出价"], ascending=False).head(300)
    app_order = bid_agg(["APP名称", "app_bundle", "订单策略素材"]).sort_values("出价", ascending=False).head(300)
    hour_order = bid_agg(["hour", "订单策略素材"]).sort_values(["hour", "出价"], ascending=[True, False]).head(300)

    filter_df = daily.copy()
    filter_df["未进入出价(估算)"] = filter_df["邀约"] - filter_df["出价"]
    filter_df["页面加载失败率"] = pct(filter_df["页面加载失败"], filter_df["落地页曝光"])
    compare_rows = []
    if not current.empty and not prev.empty:
        c = current.iloc[0]
        p = prev.iloc[0]
        for metric in ["邀约", "出价", "竞得", "点击", "落地页曝光", "跳转", "花销(元)"]:
            compare_rows.append({"指标": metric, DAY: c[metric], PREV_DAY: p[metric], "较前日变化": c[metric] - p[metric], "较前日变化率": (c[metric] - p[metric]) / p[metric] if p[metric] else 0})
    compare = pd.DataFrame(compare_rows)

    notes = pd.DataFrame(
        {
            "项目": ["漏斗定义", "花销口径", "订单限定", "直接跳转/无落地页", "时间归因", "关联键"],
            "说明": [
                "邀约、出价、竞得(广告得到曝光)、点击、落地页曝光(pageExp)、跳转(jumpButton)",
                "sum(wpr)/100000 元；wpr按分/千次曝光理解，需业务侧最终确认",
                f"当前仅分析订单 {ORDER_ID}；请求表无订单字段，邀约统计为可关联到该订单出价的请求机会。" if ORDER_ID is not None else "未限定订单，按全量投放统计。",
                "curl非空、land_page为空且无pageExp的点击后路径，漏斗到点击即可",
                "APP/小时按request_time归因；出价、竞得、点击按响应关联请求归因；APP名称由包名映射表转换",
                "request.request_id=response.response_id 且 request.imp_id=response.bid_imp_id；result/click按response_id+bid_seat_id；track优先从URL解析rid/pyck后关联，jumpButton不再使用parent_event_id直连点击",
            ],
        }
    )

    with pd.ExcelWriter(OUT, engine="openpyxl") as writer:
        add_rates(current).to_excel(writer, "大盘指标汇总", index=False)
        filter_df.to_excel(writer, "过滤情况", index=False, startrow=0)
        compare.to_excel(writer, "过滤情况", index=False, startrow=len(filter_df) + 3)
        gantt.to_excel(writer, "投放策略甘特图", index=False)
        order_fit.to_excel(writer, "订单策略创意", index=False)
        strategy_fit.to_excel(writer, "订单策略创意", index=False, startrow=len(order_fit) + 3)
        creative_fit.head(200).to_excel(writer, "订单策略创意", index=False, startrow=len(order_fit) + len(strategy_fit) + 6)
        app.to_excel(writer, "APP漏斗Top30", index=False)
        landing.to_excel(writer, "落地页情况", index=False)
        hourly.to_excel(writer, "小时分析", index=False)
        app_hour.to_excel(writer, "APP×时段", index=False)
        app_order.to_excel(writer, "APP×订单策略素材", index=False)
        hour_order.to_excel(writer, "时段×订单策略素材", index=False)
        notes.to_excel(writer, "口径说明", index=False)

        wb = writer.book
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            ws.sheet_view.showGridLines = False
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True, color="FFFFFF")
                cell.fill = cell.fill.copy(fill_type="solid", fgColor="1F4E78")
            for col in ws.columns:
                max_len = 0
                letter = col[0].column_letter
                for cell in col:
                    val = "" if cell.value is None else str(cell.value)
                    max_len = max(max_len, min(len(val), 60))
                    if isinstance(cell.value, float):
                        if "率" in str(ws.cell(1, cell.column).value) or "→" in str(ws.cell(1, cell.column).value):
                            cell.number_format = "0.00%"
                        elif "花销" in str(ws.cell(1, cell.column).value) or "CPC" in str(ws.cell(1, cell.column).value):
                            cell.number_format = "#,##0.00"
                        else:
                            cell.number_format = "#,##0.00"
                    elif isinstance(cell.value, int):
                        cell.number_format = "#,##0"
                ws.column_dimensions[letter].width = max(10, min(max_len + 2, 42))

    print(OUT.resolve())
    print("wide_rows", len(wide), "daily_rows", len(daily), "analysis_sheets", 10)


if __name__ == "__main__":
    main()
