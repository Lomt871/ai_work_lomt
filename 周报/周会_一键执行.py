import os
import traceback
from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path

from playwright.sync_api import sync_playwright


BASE_DIR = Path(__file__).resolve().parent


def load_module(filename: str, module_name: str):
    module_path = BASE_DIR / filename
    spec = spec_from_file_location(module_name, module_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法加载模块: {module_path}")
    module = module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def run_overseas_report(page, download_dir: Path, module_haiwai, module_gmv):
    print("\n" + "=" * 60)
    print("开始下载：海外团队达成")
    print("=" * 60)

    page.goto(module_haiwai.REPORT_URL)
    page.wait_for_load_state("load", timeout=60000)
    module_haiwai.time.sleep(15)

    module_gmv.export_excel(page, download_dir, module_haiwai.REPORT_FILENAME)


def main():
    mod_gmv = load_module("池内GMV目标达成_周报.py", "weekly_gmv")
    mod_haiwai = load_module("海外团队达成_周报.py", "weekly_haiwai")
    mod_yixu = load_module("一续升舱续费率_周报.py", "weekly_yixu")
    mod_tonghe = load_module("统合早鸟续费率_周报.py", "weekly_tonghe")
    mod_jinsanyue = load_module("近三月结课续费率_周报.py", "weekly_jinsanyue")
    mod_yidan = load_module("一单结课续费率（各节点涨幅）_周报.py", "weekly_yidan")
    mod_fenchi = load_module("结课月续费率_分续费池_周报.py", "weekly_fenchi")
    mod_summary = load_module("周会分析结果_生成.py", "weekly_summary")

    download_dir = mod_gmv.resolve_download_dir()
    os.environ["WEEKLY_REPORT_DOWNLOAD_DIR"] = str(download_dir)
    os.environ["WEEKLY_REPORT_MODE"] = "weekly"
    os.environ["WEEKLY_REPORT_INCLUDE_POOL_GMV"] = "1"
    os.environ["WEEKLY_REPORT_USE_MONTH_TARGET_ONLY"] = "0"

    print("\n" + "=" * 60)
    print("半周会一键执行开始")
    print("前7个脚本将复用同一次 SmartBI 登录")
    print(f"下载目录: {download_dir}")
    print("=" * 60)

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            channel="msedge",
            downloads_path=str(download_dir),
        )
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            print("\n=== 登录 SmartBI（仅一次）===")
            mod_gmv.login(page)

            print("\n1/7 池内GMV目标达成")
            page.goto(mod_gmv.REPORT_URL)
            page.wait_for_load_state("load", timeout=60000)
            mod_gmv.time.sleep(15)
            mod_gmv.set_date_to_yesterday(page)
            mod_gmv.time.sleep(5)
            mod_gmv.click_refresh(page)
            mod_gmv.export_excel(page, download_dir, mod_gmv.REPORT_FILENAME)

            print("\n2/7 海外团队达成")
            run_overseas_report(page, download_dir, mod_haiwai, mod_gmv)

            print("\n3/7 一续升舱续费率")
            page.goto(mod_yixu.REPORT_URL)
            page.wait_for_load_state("load", timeout=60000)
            mod_yixu.time.sleep(15)
            mod_yixu.set_date_to_yesterday(page)
            mod_yixu.time.sleep(5)
            mod_yixu.click_refresh(page)
            mod_yixu.export_excel(page, download_dir, mod_yixu.REPORT_FILENAME)

            print("\n4/7 统合早鸟续费率")
            page.goto(mod_tonghe.REPORT_URL)
            page.wait_for_load_state("load", timeout=60000)
            mod_tonghe.time.sleep(15)
            mod_tonghe.set_date_to_yesterday(page)
            mod_tonghe.time.sleep(5)
            mod_tonghe.set_renewal_pool_to_seven(page)
            mod_tonghe.time.sleep(5)
            mod_tonghe.click_refresh(page)
            mod_tonghe.export_excel(page, download_dir, mod_tonghe.REPORT_FILENAME)

            print("\n5/7 近三月结课续费率（2份）")
            mod_jinsanyue.open_report(page)
            mod_jinsanyue.set_date_to_yesterday(page)
            mod_jinsanyue.time.sleep(5)
            mod_jinsanyue.click_refresh(page)
            mod_jinsanyue.export_excel(page, download_dir, "一单近三月结课续费率.xlsx")

            mod_jinsanyue.open_report(page)
            mod_jinsanyue.set_date_to_yesterday(page)
            mod_jinsanyue.time.sleep(5)
            mod_jinsanyue.set_renewal_pool(page)
            mod_jinsanyue.time.sleep(5)
            mod_jinsanyue.click_refresh(page)
            mod_jinsanyue.export_excel(page, download_dir, "统合近三月结课续费率.xlsx")

            print("\n6/7 一单结课续费率（各节点涨幅，4份+汇总）")
            page.goto(mod_yidan.REPORT_URL)
            page.wait_for_load_state("load", timeout=60000)
            mod_yidan.time.sleep(15)
            mod_yidan.set_date_to_yesterday(page)
            mod_yidan.time.sleep(3)
            mod_yidan.set_logical_class_end_month(page)
            mod_yidan.time.sleep(3)
            mod_yidan.set_logical_class_end_month_end(page)
            mod_yidan.time.sleep(3)
            mod_yidan.download_report(page, download_dir, "一单结课续费率（各节点涨幅）_整体.xlsx")
            mod_yidan.download_report(page, download_dir, "一单结课续费率（各节点涨幅）_欧美澳.xlsx", "欧美澳亚海外其他国内")
            mod_yidan.download_report(page, download_dir, "一单结课续费率（各节点涨幅）_港澳.xlsx", "港澳")
            mod_yidan.download_report(page, download_dir, "一单结课续费率（各节点涨幅）_台湾.xlsx", "台湾")
            mod_yidan.merge_reports(download_dir)

            print("\n7/7 结课月续费率_分续费池（4份+汇总）")
            page.goto(mod_fenchi.REPORT_URL)
            page.wait_for_load_state("load", timeout=60000)
            mod_fenchi.time.sleep(15)
            mod_fenchi.set_date_to_yesterday(page)
            mod_fenchi.time.sleep(3)
            mod_fenchi.set_logical_class_end_month(page)
            mod_fenchi.time.sleep(3)
            mod_fenchi.set_logical_class_end_month_end(page)
            mod_fenchi.time.sleep(3)
            mod_fenchi.download_report(page, download_dir, "结课月续费率_分续费池_整体.xlsx")
            mod_fenchi.download_report(page, download_dir, "结课月续费率_分续费池_欧美澳.xlsx", "欧美澳亚海外其他国内")
            mod_fenchi.download_report(page, download_dir, "结课月续费率_分续费池_港澳.xlsx", "港澳")
            mod_fenchi.download_report(page, download_dir, "结课月续费率_分续费池_台湾.xlsx", "台湾")
            mod_fenchi.merge_reports(download_dir)

        finally:
            browser.close()

    print("\n=== 前7个下载脚本执行结束，开始执行周会分析结果生成 ===")
    try:
        mod_summary.main()
    except Exception as exc:
        print(f"周会分析结果_生成执行失败: {exc}")
        traceback.print_exc()
        raise

    print("\n全部执行完成。")


if __name__ == "__main__":
    main()
