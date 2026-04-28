"""AI 润色 — 使用可配置模型将零散笔记整理为结构化文档"""

import os

from anthropic import Anthropic

_DEFAULT_API_BASE_URL = 'https://api.aipaibox.com'
_API_BASE_URL = os.getenv('MDTONG_AI_BASE_URL', _DEFAULT_API_BASE_URL).strip() or _DEFAULT_API_BASE_URL
_MODEL = os.getenv('MDTONG_AI_MODEL', 'gpt-5.5').strip() or 'gpt-5.5'
_CONTEXT_TOKENS = int(os.getenv('MDTONG_AI_CONTEXT_TOKENS', '1000000'))
_MAX_TOKENS = int(os.getenv('MDTONG_AI_MAX_TOKENS', '128000'))
_TEMPERATURE = float(os.getenv('MDTONG_AI_TEMPERATURE', '0.2'))
_TOP_P = float(os.getenv('MDTONG_AI_TOP_P', '0.95'))
_MAX_CONTINUATIONS = 4
_INCOMPLETE_STOP_REASONS = {'max_tokens', 'model_context_window_exceeded'}
_COMPLETION_SENTINEL = '<!--MDTONG_POLISH_COMPLETE-->'
_BASE_SYSTEM_SUFFIX = f"""

完整性与质量要求：
- 当前模型上下文上限为 {_CONTEXT_TOKENS} tokens，最大输出上限为 {_MAX_TOKENS} tokens；请充分利用上下文，完整覆盖用户提供的全部信息。
- 不得因内容较长而省略、概括性截断或用“略”“待补充”等方式替代原文关键信息。
- 若内容很多，请继续输出直到文档自然结束。
- 文档末尾必须单独追加完整性标记：{_COMPLETION_SENTINEL}
""".strip()

PROMPT_TEMPLATES = {
    '会议纪要': {
        'system': """你是一位专业的会议纪要整理专员。你的任务是将用户提供的零散会议笔记整理为结构化的 Markdown 格式会议纪要。

输出格式要求：
1. 使用 `# 标题` 作为会议纪要标题
2. 使用 `**标签：** 内容` 格式列出元数据（会议时间、地点、参会人员等）
3. 使用 `## 一、二、三、` 等中文编号作为章节标题
4. 正文内容保持清晰简洁，去除口语化表达
5. 如有数据或对比信息，使用 Markdown 表格
6. 保留所有关键信息，不要遗漏重要细节
7. 如果原文已经是较好的格式，只做润色和规范化处理

直接输出整理后的 Markdown 文本，不要添加额外说明。""",
        'user': '请整理以下会议笔记：\n\n{text}',
    },
    '普通文档': {
        'system': """你是一位专业的文档编辑。你的任务是将用户提供的文本整理为清晰规范的 Markdown 文档。

输出格式要求：
1. 使用 `# 标题` 作为文档主标题
2. 使用 `##`、`###` 等层级标题组织内容结构
3. 正文语言简洁流畅，逻辑清晰
4. 适当使用列表、加粗、表格等 Markdown 元素增强可读性
5. 去除冗余和口语化表达，保留所有关键信息
6. 如果原文已经是较好的格式，只做润色和规范化处理

直接输出整理后的 Markdown 文本，不要添加额外说明。""",
        'user': '请整理以下文本为规范的 Markdown 文档：\n\n{text}',
    },
    '设计文档': {
        'system': """你是一位资深软件架构师。你的任务是将用户提供的设计思路或技术方案整理为结构化的设计文档。

输出格式要求：
1. `# 项目/模块名称 — 设计文档`
2. `## 一、背景与目标`：阐述问题背景和设计目标
3. `## 二、整体架构`：系统架构概览，关键组件说明
4. `## 三、详细设计`：核心模块的接口、数据流、状态管理等
5. `## 四、数据库/数据模型`（如适用）：表结构或数据schema
6. `## 五、接口设计`（如适用）：API 接口列表，使用 Markdown 表格
7. `## 六、非功能性需求`：性能、安全、可扩展性等
8. `## 七、风险与待定事项`
9. 使用代码块展示关键接口签名或数据结构
10. 使用表格整理接口列表、字段定义等

直接输出整理后的 Markdown 文本，不要添加额外说明。""",
        'user': '请将以下内容整理为设计文档：\n\n{text}',
    },
    '需求文档': {
        'system': """你是一位资深产品经理。你的任务是将用户提供的需求描述整理为结构化的产品需求文档（PRD）。

输出格式要求：
1. `# 产品/功能名称 — 需求文档`
2. `## 一、需求背景`：业务背景、用户痛点、市场分析
3. `## 二、需求目标`：核心目标和关键指标（KPI）
4. `## 三、用户角色`：目标用户画像
5. `## 四、功能需求`：按优先级排列，使用 `### P0/P1/P2` 分级
   - 每个功能包含：功能描述、用户故事、验收标准
6. `## 五、非功能需求`：性能、安全、兼容性要求
7. `## 六、交互流程`：核心用户操作流程
8. `## 七、数据需求`（如适用）：需要的数据支持
9. `## 八、排期与里程碑`（如有信息）
10. 使用表格整理功能清单、优先级矩阵等

直接输出整理后的 Markdown 文本，不要添加额外说明。""",
        'user': '请将以下内容整理为需求文档：\n\n{text}',
    },
    '周报': {
        'system': """你是一位高效的职场助理。你的任务是将用户提供的工作记录整理为清晰的周报。

输出格式要求：
1. `# 周报 — YYYY年第N周`（根据内容推断日期）
2. `**姓名：**`、`**部门：**`、`**日期范围：**`（如有信息）
3. `## 一、本周工作完成情况`：按项目/模块分类，列出完成的工作项
4. `## 二、进行中的工作`：当前正在推进但未完成的事项
5. `## 三、遇到的问题与风险`：阻塞项、需要协助的事项
6. `## 四、下周工作计划`：按优先级排列
7. 每个工作项简洁明了，使用列表格式
8. 如有进度百分比或数据，使用表格展示

直接输出整理后的 Markdown 文本，不要添加额外说明。""",
        'user': '请将以下工作记录整理为周报：\n\n{text}',
    },
    '公文': {
        'system': """你是一位资深的体制内公文写作专家。你的任务是将用户提供的内容严格按照《党政机关公文格式》（GB/T 9704—2012）标准整理为规范的公文格式 Markdown 文档。

你必须严格遵循以下公文结构和格式要求：

## 公文结构（自上而下）

### 版头部分
1. `# 发文机关名称`（标题居中，对应红头文件的发文机关标志）
2. 空行后用 `---` 分隔线（对应红色分隔线）
3. `**发文字号：**` 如"X办发〔2026〕X号"（机关代字＋年份＋序号，年份用六角括号〔〕）
4. 如有签发人，使用 `**签发人：**` 标注

### 标题部分
5. `## 关于XXX的通知/通报/请示/批复/函/意见/报告/决定`
   - 标题结构："关于" + 事由 + 文种
   - 标题应概括准确、简洁明了

### 主送机关
6. `**主送：**`（顶格写，后加冒号，列出主送机关全称或规范化简称）

### 正文部分
7. 正文分为以下层次，严格使用以下编号体系（不得跳级、混用）：
   - 第一层：`## 一、`（用中文数字加顿号）
   - 第二层：`**（一）**`（用中文数字加括号，括号外无标点）
   - 第三层：`1.`（用阿拉伯数字加下脚点）
   - 第四层：`（1）`（用阿拉伯数字加括号）
8. 正文语言要求：
   - 使用规范的公文用语，庄重严谨
   - 禁止口语化、网络用语
   - 常用公文用语示例：兹、拟、特此通知、现将有关事项通知如下、请遵照执行、妥否请批示
   - 如涉及时间，使用"XXXX年X月X日"全称格式
   - 数字使用：序号和编号用中文或阿拉伯数字（按层级），其他计量数据用阿拉伯数字

### 附件说明（如有）
9. 正文之后另起一行写 `**附件：**`，列出附件名称

### 落款部分
10. `**发文机关署名：**`（写全称或规范化简称）
11. `**成文日期：**`（XXXX年X月X日，用阿拉伯数字）

### 版记部分（如有）
12. 用 `---` 分隔线隔开
13. `**抄送：**`（抄送机关，用顿号分隔）
14. `**印发机关及日期：**`

## 常见公文文种及注意事项
- **通知**：告知性公文，结尾用"特此通知"
- **通报**：表彰、批评或传达情况，结尾用"特此通报"
- **请示**：下级向上级请求指示或批准，一文一事，结尾用"妥否，请批示"或"以上请示，请予审批"
- **批复**：上级答复下级的请示，开头引述来文标题和发文字号
- **报告**：向上级汇报工作、反映情况，结尾用"特此报告"
- **函**：平行或不相隶属机关之间商洽工作，语气相对灵活
- **意见**：对重要问题提出见解和处理办法
- **决定**：对重要事项作出决策和部署

## 重要规范
- 如果用户未提供发文机关、文号等信息，应根据内容合理推断或使用"XX机关""XX部门"等占位
- 如果用户未指定文种，根据内容性质判断最合适的文种
- 结构层次编号严格按规范使用，不得使用"第一""首先"等非标准编号
- 每段段首不缩进（Markdown 不需要首行缩进，DOCX 生成时由排版系统处理）

直接输出整理后的 Markdown 文本，不要添加额外说明。""",
        'user': '请将以下内容整理为标准公文格式：\n\n{text}',
    },
}


def get_template_names():
    return list(PROMPT_TEMPLATES.keys())


def _resolve_api_key(api_key=None):
    """按调用参数、专用环境变量、Anthropic 环境变量的顺序解析 API Key。"""
    resolved = (api_key or os.getenv('MDTONG_AI_API_KEY') or os.getenv('ANTHROPIC_API_KEY') or '').strip()
    if not resolved:
        raise ValueError('未配置 AI API Key：请在设置中填写 API Key，或设置 MDTONG_AI_API_KEY/ANTHROPIC_API_KEY 环境变量')
    return resolved


def _create_client(api_key=None, base_url=None):
    return Anthropic(api_key=_resolve_api_key(api_key), base_url=base_url or _API_BASE_URL)


def _extract_usage(resp_usage):
    return {
        'input_tokens': getattr(resp_usage, 'input_tokens', 0),
        'output_tokens': getattr(resp_usage, 'output_tokens', 0),
        'cache_read': getattr(resp_usage, 'cache_read_input_tokens', 0),
        'cache_creation': getattr(resp_usage, 'cache_creation_input_tokens', 0),
    }


def _build_system_prompt(system_text):
    return f'{system_text}\n\n{_BASE_SYSTEM_SUFFIX}'


def _strip_completion_sentinel(text):
    return text.replace(_COMPLETION_SENTINEL, '').rstrip()


def _is_complete(text, stop_reason):
    if _COMPLETION_SENTINEL in text:
        return True
    return stop_reason not in _INCOMPLETE_STOP_REASONS


def _merge_usage(total, usage):
    for key in ('input_tokens', 'output_tokens', 'cache_read', 'cache_creation'):
        total[key] = total.get(key, 0) + usage.get(key, 0)
    return total


def _extract_text(response):
    parts = []
    for block in response.content:
        if block.type == 'text':
            parts.append(block.text)
    return ''.join(parts)


def _build_messages(user_msg, partial_text=None):
    if not partial_text:
        return [{'role': 'user', 'content': user_msg}]
    continuation_prompt = f"""上一次输出因长度限制可能未完整结束。请严格从下方“已生成内容”的末尾继续生成，不要重复已生成内容，不要重新开始。

已生成内容：
{partial_text}

续写要求：
1. 只输出后续缺失的 Markdown 内容。
2. 保持章节编号、语气和格式连续。
3. 完成全部内容后，在末尾单独追加 {_COMPLETION_SENTINEL}。"""
    return [
        {'role': 'user', 'content': user_msg},
        {'role': 'assistant', 'content': partial_text},
        {'role': 'user', 'content': continuation_prompt},
    ]


def polish(raw_text, api_key=None, model=None, template='会议纪要', on_chunk=None):
    """
    调用可配置模型润色文档。

    Args:
        raw_text: 原始文本
        api_key: API Key，未传入时读取 MDTONG_AI_API_KEY 或 ANTHROPIC_API_KEY
        model: 模型名称，默认使用 gpt-5.5 或 MDTONG_AI_MODEL
        template: 提示词模板名称
        on_chunk: 流式回调函数，接收每个文本片段
    Returns:
        (润色后文本, usage_dict) 二元组
    """
    tpl = PROMPT_TEMPLATES.get(template, PROMPT_TEMPLATES['会议纪要'])
    system_text = _build_system_prompt(tpl['system'])
    user_msg = tpl['user'].format(text=raw_text)

    system_with_cache = [
        {
            'type': 'text',
            'text': system_text,
            'cache_control': {'type': 'ephemeral'},
        }
    ]

    client = _create_client(api_key=api_key)
    selected_model = model or _MODEL
    usage_total = {}
    result = []
    partial_text = ''

    for attempt in range(_MAX_CONTINUATIONS + 1):
        common_params = {
            'model': selected_model,
            'max_tokens': _MAX_TOKENS,
            'temperature': _TEMPERATURE,
            'top_p': _TOP_P,
            'system': system_with_cache,
            'messages': _build_messages(user_msg, partial_text if attempt else None),
        }

        if on_chunk:
            chunk_parts = []
            sentinel_seen = False
            with client.messages.stream(**common_params) as stream:
                for text in stream.text_stream:
                    chunk_parts.append(text)
                    if sentinel_seen:
                        continue
                    if _COMPLETION_SENTINEL in text:
                        before_sentinel = text.split(_COMPLETION_SENTINEL, 1)[0]
                        if before_sentinel:
                            on_chunk(before_sentinel)
                        sentinel_seen = True
                    else:
                        on_chunk(text)
                final = stream.get_final_message()
            chunk_text = ''.join(chunk_parts)
            stop_reason = getattr(final, 'stop_reason', None) if final else None
            usage = _extract_usage(final.usage) if final else {}
        else:
            response = client.messages.create(**common_params)
            chunk_text = _extract_text(response)
            stop_reason = getattr(response, 'stop_reason', None)
            usage = _extract_usage(response.usage)

        usage_total = _merge_usage(usage_total, usage)
        result.append(chunk_text)
        partial_text = ''.join(result)

        if _is_complete(partial_text, stop_reason):
            polished = _strip_completion_sentinel(partial_text)
            return polished, usage_total

    polished = _strip_completion_sentinel(partial_text)
    return polished, usage_total


import re as _re

_DOC_TYPE_RULES = [
    ('会议纪要', [
        _re.compile(r'会议|纪要|参会|与会|出席|议题|会上|讨论|决议|会议室', _re.IGNORECASE),
    ]),
    ('周报', [
        _re.compile(r'周报|本周|下周|上周|工作总结|工作计划|完成情况|进行中', _re.IGNORECASE),
    ]),
    ('公文', [
        _re.compile(r'通知|通报|请示|批复|函|意见|报告|决定|发文|签发|各[单部处科室]|特此|遵照执行|关于.*的(通知|请示|报告|意见)', _re.IGNORECASE),
    ]),
    ('需求文档', [
        _re.compile(r'需求|PRD|用户故事|用户角色|功能需求|验收标准|产品需求|用例|场景', _re.IGNORECASE),
    ]),
    ('设计文档', [
        _re.compile(r'设计文档|架构|技术方案|接口设计|数据模型|系统设计|模块设计|API|数据库设计|ER图', _re.IGNORECASE),
    ]),
]


def detect_doc_type(text):
    """根据文本内容关键词自动推荐文档模板类型"""
    if not text or len(text.strip()) < 10:
        return '普通文档'

    sample = text[:2000]
    scores = {}

    for tpl_name, patterns in _DOC_TYPE_RULES:
        score = 0
        for pat in patterns:
            matches = pat.findall(sample)
            score += len(matches)
        if score > 0:
            scores[tpl_name] = score

    if scores:
        return max(scores, key=scores.get)
    return '普通文档'
