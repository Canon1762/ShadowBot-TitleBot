# -*- coding: utf-8 -*-
"""
影刀 RPA 6.3.22 社区版（64 位）专用：Gemini 商品标题写入 WPS

用法：
1. 在影刀 6.3.22 的可视化流程里，搜索“插入代码段(Python)”；
2. 把本文件从第一行到最后一行整段粘贴进去；
3. 只修改下面“只改这里”的 10 个设置；
4. 先令 MAX_ITEMS = 1 测试一行，成功后再改成正式数量。

这段代码不调用任何 AI API，不使用 OCR，不依赖屏幕坐标。
"""

import json
import re
import time

# 影刀的代码段环境已经自动导入 xbot；这里显式导入子模块，
# 避免个别 6.3.22 小版本需要手动加载 xbot.web / xbot.excel。
# 如果你的影刀报 import xbot.web 失败，把下面两行删掉，
# 并改用教程里的“积木版”绑定网页和 Excel，代码段只做字符串处理。
import xbot.web
import xbot.excel


# ==================== 只改这里 ====================
TARGET_COL = "B"                  # 目标列，例如 B
START_ROW = 2                     # 第一个绿色行；如果第一个绿色行是 86，就填 86
ROW_STEP = 4                      # 每次向下跳几行；你的表是 4
MAX_ITEMS = 10                    # 先填 1 测试；正式批量时填实际条数
SHEET_NAME = ""                   # 留空 = 当前激活的工作表；也可以填 "Sheet1"
GEMINI_URL_KEYWORD = "gemini.google.com"
BROWSER_MODE = "chrome"           # chrome=谷歌浏览器；edge=Edge；cef=影刀浏览器
PAGE_PICK_MODE = "active"          # active=只抓当前选中的网页；latest=多个匹配时取最新打开的网页
READ_MODE = "batch"               # single=每次读最后一条标题；batch=一次读最后回复里的所有标题
WAIT_BEFORE_READ_SECONDS = 0      # Gemini 还在生成时，可以改成 1 或 2
SAVE_AFTER_RUN = True             # True=跑完后自动保存 WPS；第一次测试保持 False
# ==================================================


class TitleError(ValueError):
    """标题清洗或替换失败。"""


_TITLE_MARKER_RE = re.compile(r"商品标题\s*[:：]\s*")
_MARKDOWN_PREFIX_RE = re.compile(r"^(?:(?:[-*•]\s*)|\d+[.、]\s*)+")
_STOP_HEADING_RE = re.compile(
    r"^(?:设计|创意|推荐|完整|英文|参数|提示词|描述|风格|用途|注意事项|注意)[^：:]*[:：]"
)


def _normalize_newlines(value):
    text = str(value)
    text = text.replace("\ufeff", "").replace("\xa0", " ")
    return text.replace("\r\n", "\n").replace("\r", "\n")


def _strip_markdown_prefix(line):
    line = _MARKDOWN_PREFIX_RE.sub("", line.strip()).strip()
    return line.strip("*` ").strip()


def _extract_after_marker(text):
    match = _TITLE_MARKER_RE.search(text)
    if not match:
        return None

    tail = text[match.end():].strip()
    if not tail:
        raise TitleError("商品标题标记后没有正文")

    picked = []
    for raw_line in tail.split("\n"):
        line = _strip_markdown_prefix(raw_line)
        if not line:
            if picked:
                break
            continue
        if picked and _STOP_HEADING_RE.match(line):
            break
        picked.append(line)
        if len(" ".join(picked)) > 300:
            break

    if not picked:
        raise TitleError("商品标题正文为空")
    return " ".join(picked).strip()


def extract_gemini_title(raw_title):
    if raw_title is None:
        raise TitleError("Gemini 读取结果为空")

    text = _normalize_newlines(raw_title).strip()
    if not text:
        raise TitleError("Gemini 读取结果为空")

    extracted = _extract_after_marker(text)
    if extracted is not None:
        title = extracted
    else:
        non_empty_lines = [line.strip() for line in text.split("\n") if line.strip()]
        if len(non_empty_lines) != 1:
            raise TitleError("未找到“商品标题：”标记，且文本是多行，已停止")
        title = _strip_markdown_prefix(non_empty_lines[0])
        if len(title) > 200:
            raise TitleError("未找到“商品标题：”标记，且文本超过 200 字，已停止")

    title = re.sub(r"\s+", " ", title).strip()
    if not title:
        raise TitleError("清洗后标题为空")
    return title


def replace_title_between_commas(old_text, new_title, row_no=None):
    row_prefix = "" if row_no is None else f"第{row_no}行 "

    if old_text is None or str(old_text).strip() == "":
        raise TitleError(f"{row_prefix}WPS 当前单元格为空")

    old = str(old_text)
    title = extract_gemini_title(new_title)

    comma_count = old.count("，")
    if comma_count < 2:
        raise TitleError(f"{row_prefix}没有找到两个中文逗号，实际找到 {comma_count} 个")

    adjacent_pair = old.find("，，")
    if adjacent_pair < 0:
        raise TitleError(f"{row_prefix}没有找到连续的中文逗号占位符“，，”，已停止")

    first_comma = adjacent_pair
    second_comma = adjacent_pair + 1
    prefix = old[:first_comma + 1]
    suffix = old[second_comma:]
    result_text = prefix + title + suffix

    if not result_text.startswith(prefix) or not result_text.endswith(suffix):
        raise TitleError(f"{row_prefix}字符串拼接自检失败")
    return result_text


JS_GET_ONE_TITLE = r"""
function () {
  const marker = /商品标题\s*[:：]\s*/;
  const responses = Array.from(document.querySelectorAll("model-response"));
  const scope = responses.length
    ? responses[responses.length - 1]
    : (document.querySelector("main") || document);

  const candidates = Array.from(
    scope.querySelectorAll("p, div, li, span, strong")
  ).filter((el) => {
    if (!el.getClientRects().length) return false;
    const text = (el.innerText || el.textContent || "")
      .replace(/\u00a0/g, " ")
      .trim();
    if (!marker.test(text)) return false;
    if (!text.replace(marker, "").trim()) return false;
    const childHasFullText = Array.from(el.children).some((child) => {
      const childText = (child.innerText || child.textContent || "")
        .replace(/\u00a0/g, " ")
        .trim();
      return marker.test(childText) && childText.replace(marker, "").trim();
    });
    return !childHasFullText;
  });

  if (!candidates.length) return "";
  const el = candidates[candidates.length - 1];
  return (el.innerText || el.textContent || "").trim();
}
"""


JS_GET_ALL_TITLES = r"""
function () {
  const marker = /商品标题\s*[:：]\s*/;
  const responses = Array.from(document.querySelectorAll("model-response"));
  const scope = responses.length
    ? responses[responses.length - 1]
    : (document.querySelector("main") || document);

  const candidates = Array.from(
    scope.querySelectorAll("p, div, li, span, strong")
  ).filter((el) => {
    if (!el.getClientRects().length) return false;
    const text = (el.innerText || el.textContent || "")
      .replace(/\u00a0/g, " ")
      .trim();
    if (!marker.test(text)) return false;
    if (!text.replace(marker, "").trim()) return false;
    const childHasFullText = Array.from(el.children).some((child) => {
      const childText = (child.innerText || child.textContent || "")
        .replace(/\u00a0/g, " ")
        .trim();
      return marker.test(childText) && childText.replace(marker, "").trim();
    });
    return !childHasFullText;
  }).map((el) => (el.innerText || el.textContent || "").trim());

  return Array.from(new Set(candidates));
}
"""


def _read_titles(browser, all_titles):
    code = JS_GET_ALL_TITLES if all_titles else JS_GET_ONE_TITLE
    result = browser.execute_javascript(code)

    if result is None:
        return []

    if all_titles:
        if isinstance(result, str):
            text = result.strip()
            if not text:
                return []
            try:
                parsed = json.loads(text)
            except Exception:
                parsed = [text]
            result = parsed
        if not isinstance(result, (list, tuple)):
            result = [result]
        return [str(item).strip() for item in result if str(item).strip()]

    text = str(result).strip()
    return [text] if text else []

def _get_target_browser():
    """多开 Gemini 页面时，默认只抓当前选中的那个网页。"""
    if PAGE_PICK_MODE == "active":
        browser = xbot.web.get_active(mode=BROWSER_MODE, load_timeout=0)

        current_url = ""
        try:
            current_url = browser.get_url() or ""
        except Exception:
            current_url = ""

        if GEMINI_URL_KEYWORD not in current_url:
            raise TitleError(
                f"当前选中的网页不是 Gemini：{current_url!r}。"
                f"请先点到要处理的 Gemini 标签页，"
                f"或把 PAGE_PICK_MODE 改成 'latest'。"
            )
        return browser

    if PAGE_PICK_MODE == "latest":
        return xbot.web.get(
            url=GEMINI_URL_KEYWORD,
            mode=BROWSER_MODE,
            load_timeout=0,
        )

    raise TitleError('PAGE_PICK_MODE 只能是 "active" 或 "latest"')

def run():
    if MAX_ITEMS <= 0:
        raise TitleError("MAX_ITEMS 必须大于 0")
    if READ_MODE not in ("single", "batch"):
        raise TitleError('READ_MODE 只能是 "single" 或 "batch"')

    browser = _get_target_browser()

    try:
        print(f"本次处理网页：{browser.get_title()} | {browser.get_url()}")
    except Exception:
        pass

    workbook = xbot.excel.get_active_workbook()
    if SHEET_NAME:
        sheet = workbook.get_sheet_by_name(SHEET_NAME)
    else:
        sheet = workbook.get_active_sheet()

    if READ_MODE == "batch":
        titles = _read_titles(browser, all_titles=True)
        if not titles:
            raise TitleError("Gemini 页面没有读取到任何“商品标题：”内容")
        if len(titles) < MAX_ITEMS:
            print(
                f"提示：Gemini 只读到 {len(titles)} 条标题，"
                f"MAX_ITEMS={MAX_ITEMS}，本次只处理 {len(titles)} 行。"
            )
        max_items = min(MAX_ITEMS, len(titles))
    else:
        titles = []
        max_items = MAX_ITEMS

    success_count = 0
    for index in range(max_items):
        row = START_ROW + index * ROW_STEP

        if WAIT_BEFORE_READ_SECONDS > 0:
            time.sleep(WAIT_BEFORE_READ_SECONDS)

        if READ_MODE == "batch":
            raw_title = titles[index]
        else:
            one = _read_titles(browser, all_titles=False)
            if not one:
                raise TitleError(f"第{row}行：Gemini 没有读取到“商品标题：”内容")
            raw_title = one[0]

        old_text = sheet.get_cell(row, TARGET_COL, using_text=True)
        result_text = replace_title_between_commas(old_text, raw_title, row_no=row)

        sheet.set_cell(row, TARGET_COL, result_text)
        verify_text = sheet.get_cell(row, TARGET_COL, using_text=True)

        if str(verify_text).strip() != str(result_text).strip():
            raise TitleError(
                f"第{row}行：回读校验不通过，"
                f"写入={result_text!r}，回读={verify_text!r}"
            )

        success_count += 1
        print(f"[第{row}行] 成功：{result_text}")

    if SAVE_AFTER_RUN:
        workbook.save()

    print(f"完成：成功 {success_count} 行。")


run()
