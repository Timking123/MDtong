#!/usr/bin/env python3
"""样本文档生成器 — 调用 AI 润色提示词生成 6 种文档类型的样本 Markdown"""

import argparse
import json
import os
import sys
import time

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, 'config.json')
SAMPLES_DIR = os.path.join(SCRIPT_DIR, 'samples')

FORMAT_HINT = """

补充要求：请在整理时充分运用各种 Markdown 格式来增强文档的可读性和专业性，包括但不限于：
- 使用表格整理对比数据和结构化信息
- 使用围栏代码块（```语言名）展示技术内容、配置、命令等
- 使用任务列表（- [x] / - [ ]）展示已完成和待办事项
- 使用有序列表和无序列表组织步骤和要点
- 使用加粗（**文本**）突出关键信息
- 使用行内代码（`代码`）标记技术术语、命令、文件名等
- 使用超链接（[文本](URL)）标注参考资料
- 使用水平线（---）分隔不同部分
- 使用引用（> 文本）标注重要提示或引述
"""

SAMPLE_INPUTS = {
    '会议纪要': """\
产品发布会议 2026年4月18号下午2点到4点半 在总部大楼3层A会议室开的
参会人：张明（产品总监，主持）、李芳（技术负责人）、王强（市场经理）、赵雪（UI设计师）、陈磊（后端开发）、刘洋（测试主管）
列席：财务部 孙悦

主要讨论了智能助手 App v2.0 的发布计划

一个是发布时间，大家讨论了很久，最终定在5月15号上线，之前5月1号要完成内测，5月8号提交应用商店审核

技术方面，李芳说新版本核心功能是接入大模型API，用的是RESTful接口，endpoint是 POST /api/v2/chat/completions，请求体格式是 {"model": "gpt-4", "messages": [{"role": "user", "content": "你好"}], "temperature": 0.7, "max_tokens": 2048}，响应格式是 {"id": "chatcmpl-xxx", "choices": [{"message": {"content": "回复内容"}}], "usage": {"prompt_tokens": 10, "completion_tokens": 20}}
目前API响应延迟P50是800ms，P99是2.3秒，目标优化到P50 500ms以内
参考技术文档 https://docs.example.com/api/v2 和架构图 https://wiki.internal.com/arch-diagram

预算情况，孙悦汇报了：
- 服务器成本：Q1花了12万，Q2预算15万，Q3预计需要20万（用户增长导致）
- API调用费用：Q1花了8万，Q2预算10万，Q3预计12万
- 人力成本：Q1是45万，Q2是48万，Q3持平48万
- 营销推广费用：Q1花了5万，Q2预算20万（因为要配合发布），Q3预计15万
总预算 Q1=70万，Q2=93万，Q3=95万

市场方面，王强说竞品分析了三个：
智能星 - 用户量500万，月活200万，主打功能是语音交互，价格免费+会员39元/月
AI伙伴 - 用户量800万，月活350万，主打功能是知识库问答，价格免费+会员49元/月
聊天宝 - 用户量300万，月活100万，主打功能是多模态，价格29元/月
我们的定价策略建议基础版免费，专业版29元/月，企业版99元/月

设计方面，赵雪展示了新UI改版方案，主要变化是采用了毛玻璃效果的深色主题，图标改用线性风格，字体从14px改为15px提升可读性

接下来要做的事情：
张明负责：完成产品需求文档终版，截止4月22号（已完成70%了）
李芳负责：完成API接入联调，截止4月30号
陈磊负责：完成后端性能优化，P99延迟降到1.5秒以内，截止5月5号
赵雪负责：输出UI切图和标注文件，截止4月25号
王强负责：制定营销推广方案并对接渠道，截止5月1号
刘洋负责：编写测试用例并完成自动化测试脚本，截止5月3号（要用pytest框架，覆盖率目标90%）
孙悦负责：Q2预算审批流程，截止4月25号

下次会议定在4月25号同一时间同一地点

补充：刘洋提到目前已知的bug有15个，其中P0级别2个，P1级别5个，P2级别8个，详见bug跟踪系统 https://jira.example.com/project/AIAPP
""",

    '普通文档': """\
Python项目开发规范与最佳实践指南 给团队新人用的

第一部分 环境搭建
推荐用 pyenv 管理Python版本，安装命令是 curl https://pyenv.run | bash，然后在 .bashrc 里加上 export PATH="$HOME/.pyenv/bin:$PATH" 和 eval "$(pyenv init -)"
项目统一用Python 3.11+，虚拟环境用 python -m venv .venv 创建
依赖管理用 pip + requirements.txt，锁版本用 pip freeze > requirements.lock
IDE推荐 VS Code，必装插件：Python、Pylint、Black Formatter、GitLens

第二部分 代码规范
命名规范：变量和函数用snake_case，类用PascalCase，常量用UPPER_SNAKE_CASE，私有属性前缀单下划线_private
类型注解是必须的，函数签名示例：def process_data(items: list[dict[str, Any]], threshold: float = 0.5) -> tuple[int, list[str]]:
文档字符串用Google风格，示例：
def calculate_score(raw_data, weights):
    \"""计算加权评分。

    Args:
        raw_data: 原始数据列表，每个元素为float
        weights: 权重字典，键为类别名，值为权重系数

    Returns:
        归一化后的评分，范围0-100

    Raises:
        ValueError: 当raw_data为空时抛出
    \"""
    pass

第三部分 项目结构
标准项目目录结构是 src/包名/ 下放主代码，tests/ 放测试，docs/ 放文档，scripts/ 放工具脚本，根目录放 pyproject.toml、README.md、.gitignore 等

第四部分 测试
测试框架用 pytest，运行命令 pytest -v --cov=src --cov-report=html
测试文件命名 test_模块名.py，测试函数命名 test_功能描述
单元测试覆盖率要求85%以上，集成测试覆盖核心流程
mock外部依赖用 unittest.mock 或 pytest-mock
一个测试示例：
import pytest
from myapp.calculator import add, divide

def test_add_positive():
    assert add(2, 3) == 5

def test_divide_by_zero():
    with pytest.raises(ZeroDivisionError):
        divide(10, 0)

第五部分 常用工具对比
代码格式化：Black vs autopep8 vs yapf
Black 零配置开箱即用 格式统一 速度快 不可自定义规则 Star数35k
autopep8 遵循PEP8 可自定义 速度中等 格式不够统一 Star数4k
yapf Google出品 高度可配置 速度慢 配置复杂 Star数14k
团队推荐用Black，行长度设置为120字符，配置放在 pyproject.toml 里

静态检查：推荐 ruff，比pylint快10-100倍，配置也在 pyproject.toml 里
pre-commit 钩子配置用 .pre-commit-config.yaml，确保提交前自动检查

性能优化tips：
大列表操作用列表推导式替代for循环（快约30%），频繁字符串拼接用 "".join() 替代 +=（快约50%），缓存重复计算用 functools.lru_cache，并发IO用 asyncio 或 concurrent.futures

参考链接：
Python官方文档 https://docs.python.org/3/
PEP 8风格指南 https://peps.python.org/pep-0008/
Real Python教程 https://realpython.com/
Google Python风格指南 https://google.github.io/styleguide/pyguide.html

附：新人入职检查清单
配置开发环境 pyenv + venv （应该已经完成了）
安装并配置IDE插件
clone项目仓库并运行 pip install -e ".[dev]"
跑通全部测试
阅读项目README和架构文档
完成第一个good-first-issue
""",

    '设计文档': """\
电商订单系统微服务化改造设计 v1.0

背景：现有单体应用随业务增长遇到瓶颈，日订单量从10万增长到80万，高峰期QPS达到5000+，数据库连接池经常打满，部署一次需要停服30分钟。需要拆分为微服务架构提升可扩展性和部署效率。

目标：1）系统可用性从99.5%提升到99.99% 2）支持日订单量200万 3）部署时间从30分钟降到5分钟 4）各服务可独立扩缩容

整体架构就是拆成几个核心服务：
订单服务 order-service：负责订单创建、查询、状态流转，技术栈 Java Spring Boot 3.x + MySQL 8.0，端口8081
库存服务 inventory-service：负责库存扣减、预占、释放，技术栈 Go 1.21 + Redis 7.x，端口8082
支付服务 payment-service：负责支付对接、退款处理，技术栈 Java Spring Boot 3.x + PostgreSQL 15，端口8083
通知服务 notification-service：负责短信/邮件/推送通知，技术栈 Python FastAPI + RabbitMQ，端口8084
用户服务 user-service：负责用户信息、地址管理，技术栈 Java Spring Boot 3.x + MySQL 8.0，端口8085
网关层用 Kong API Gateway，服务发现用 Consul，配置中心用 Nacos

核心接口设计：
POST /api/orders 创建订单 参数: user_id, items[], address_id, coupon_code 返回: order_id, total_amount, status
GET /api/orders/{id} 查询订单 参数: 无 返回: 完整订单信息含商品列表
PUT /api/orders/{id}/cancel 取消订单 参数: reason 返回: refund_id, status
POST /api/orders/{id}/pay 支付订单 参数: payment_method, payment_token 返回: transaction_id, status
GET /api/orders?user_id=X&status=Y&page=1&size=20 订单列表 返回: 分页订单列表

数据库设计，订单表核心字段：
CREATE TABLE orders (
    id BIGINT PRIMARY KEY AUTO_INCREMENT,
    order_no VARCHAR(32) NOT NULL UNIQUE COMMENT '订单编号',
    user_id BIGINT NOT NULL COMMENT '用户ID',
    total_amount DECIMAL(10,2) NOT NULL COMMENT '订单总金额',
    discount_amount DECIMAL(10,2) DEFAULT 0 COMMENT '优惠金额',
    pay_amount DECIMAL(10,2) NOT NULL COMMENT '实付金额',
    status TINYINT NOT NULL DEFAULT 0 COMMENT '0-待支付 1-已支付 2-已发货 3-已完成 4-已取消 5-已退款',
    payment_method VARCHAR(20) COMMENT '支付方式',
    address_snapshot JSON COMMENT '地址快照',
    remark VARCHAR(500) COMMENT '备注',
    created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    INDEX idx_user_id (user_id),
    INDEX idx_order_no (order_no),
    INDEX idx_status_created (status, created_at)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

订单状态流转：待支付→已支付→已发货→已完成，待支付→已取消，已支付→已退款

消息队列使用RabbitMQ，Exchange用 order.topic 类型TopicExchange
routing key 设计：order.created（新建通知库存扣减）、order.paid（支付成功通知发货）、order.cancelled（取消通知库存释放）、order.completed（完成通知积分奖励）

非功能需求：
接口响应时间P99 < 200ms，吞吐量 > 10000 QPS
数据一致性用Saga模式处理分布式事务
服务间通信用gRPC（内部）+ REST（外部）
监控用 Prometheus + Grafana，链路追踪用 Jaeger
日志用 ELK Stack（Elasticsearch + Logstash + Kibana）
参考架构文档 https://microservices.io/patterns

风险和待定事项：
数据迁移方案还没定，预计需要2周
旧系统兼容期多长还在讨论
Go 语言库存服务的团队培训还没安排
""",

    '需求文档': """\
App 消息推送中心 需求描述 产品经理：李明月 最后更新 2026-04-15

背景说明：现在app的推送功能很零散，促销推送、订单通知、系统消息各走各的通道，用户经常反馈收不到重要通知或者被营销推送打扰。客户满意度调研中，通知体验评分只有3.2/5分。竞品分析发现头部App都有统一的消息中心。本次需求要建设统一的消息推送中心，提升用户体验和消息送达率。

目标用户：
C端普通用户 - 需要接收订单状态、活动优惠、系统公告等消息，年龄18-45岁，日均使用App 25分钟
商家用户 - 需要接收订单提醒、经营数据通知、平台公告，日均使用App 40分钟
运营人员 - 需要配置推送策略、查看推送数据、管理消息模板

核心KPI：消息送达率从75%提升到95%，用户通知设置页访问后投诉率下降50%，消息中心日活渗透率达到30%

功能需求列表：

P0 必须有的：
1. 统一消息中心页面 - 用户在App首页可以看到未读消息数角标，点进去看到按时间排列的所有消息，分为交易消息、活动消息、系统通知三个tab。用户故事：作为用户，我想在一个地方查看所有消息，这样就不会遗漏重要通知。验收标准：消息按时间倒序展示，支持下拉刷新和上滑加载更多，未读消息有视觉标记，支持全部已读操作
2. 推送偏好设置 - 用户可以按消息类型开关推送通知（交易消息默认开且不可关，活动消息默认开可关，系统通知默认开可关），可以设置免打扰时段。验收标准：设置实时生效，服务端记录用户偏好，免打扰时段内消息静默但仍可在消息中心查看
3. 消息模板管理后台 - 运营人员可以创建和管理推送消息模板，支持变量替换（如{用户名}、{订单号}），预览效果。验收标准：模板支持标题+正文+图片+跳转链接，变量支持自动校验

P1 应该有的：
4. 智能推送策略 - 根据用户活跃时段智能选择推送时机，A/B测试不同推送文案的效果。验收标准：推送时机优化后点击率提升15%
5. 消息已读回执 - 记录消息的送达、展示、点击等状态。验收标准：埋点数据准确率99%以上
6. 富媒体消息 - 支持图文混排、按钮操作（如直接在推送中完成评价）

P2 可以后做的：
7. 消息搜索 - 在消息中心搜索历史消息
8. 消息收藏 - 用户可以收藏重要消息
9. 定时推送 - 运营可以设置定时发送的推送任务

交互流程（核心-接收推送并查看详情）：
第一步 服务端触发推送事件（如订单发货）
第二步 推送引擎检查用户偏好和免打扰设置
第三步 根据用户设备类型（iOS/Android）调用对应推送通道（APNs/FCM）
第四步 用户设备展示推送通知（标题+摘要）
第五步 用户点击推送通知，App打开对应详情页
第六步 标记消息已读，上报点击事件

数据需求：
需要的数据表和字段大概有：
消息记录表 - 消息ID、用户ID、消息类型（1交易2活动3系统）、标题、正文、图片URL、跳转URL、发送时间、送达时间、阅读时间、状态（0未读1已读2已删）
用户推送偏好表 - 用户ID、交易消息开关、活动消息开关、系统通知开关、免打扰开始时间、免打扰结束时间、更新时间
推送任务表 - 任务ID、任务名称、消息模板ID、目标用户群、推送时间、状态（0待发送1发送中2已完成3已取消）、发送量、送达量、点击量

技术约束：
推送SDK用极光推送 jpush-api-python-client，文档见 https://docs.jiguang.cn/jpush/
单次推送目标用户上限100万，超过需要分批
消息存储保留90天，过期自动归档到冷存储
推送频率限制：单用户每天最多接收10条营销推送

排期计划：
需求评审 4月20号，技术方案评审 4月25号，P0功能开发 4月28号到5月20号（3周），P1功能开发 5月21号到6月5号（2周），联调测试 6月6号到6月15号（1.5周），灰度发布 6月16号，全量上线 6月20号
""",

    '周报': """\
陈磊 后端开发组 2026年4月14日-4月18日的工作记录

本周完成的事情：
1. 订单查询接口性能优化，给orders表加了联合索引 (user_id, status, created_at)，查询耗时从平均320ms降到45ms，提升了86%。修改了SQL查询，原来写的是 SELECT * FROM orders WHERE user_id=? ORDER BY created_at DESC，优化后改成 SELECT id,order_no,status,pay_amount,created_at FROM orders WHERE user_id=? AND status IN (1,2,3) ORDER BY created_at DESC LIMIT 20，减少了不必要的全字段查询
2. 修复了用户头像上传的OOM bug，原因是没有限制图片大小就直接读进内存，修复方案是加了文件大小校验（最大5MB）和图片压缩处理，用Pillow库做resize，代码大概是 img = Image.open(file); if max(img.size) > 1920: img.thumbnail((1920, 1920)); img.save(output, quality=85)
3. 完成了支付回调接口的幂等性改造，用Redis的SETNX实现了分布式锁，锁的key格式是 payment:callback:{transaction_id}，TTL设置为10分钟
4. 参与了新版本的Code Review，提了12个MR评审意见，合入了8个MR
5. 整理了API文档，更新了Swagger注释

正在做还没做完的：
Redis缓存层改造，大概完成了60%，把热点数据（商品详情、用户基本信息）从MySQL查询改成先查Redis。预计下周三完成。遇到的问题是缓存更新策略在并发场景下可能有短暂不一致，正在调研Cache Aside Pattern和Write Through的最佳实践
消息队列RabbitMQ的死信队列配置也在搞，完成了40%

遇到的问题：
staging环境的MySQL 8.0有个bug，在大事务回滚时偶尔会出现死锁，已报告给DBA团队，临时方案是把大事务拆成了小批量操作。相关issue https://github.com/company/backend/issues/234
还有就是测试环境的Redis集群上周五挂了一次，导致相关功能测试延迟了半天

这周的一些数据：
项目 上周 本周 变化
API平均响应时间 185ms 142ms 下降23%
错误率 0.15% 0.08% 下降47%
代码覆盖率 76% 81% 提升5个百分点
已解决Bug数 6个 9个 增加50%
新增代码行数 约2400行 约1800行 减少（主要是重构删了冗余代码）

下周计划做的事：
完成Redis缓存层改造的剩余部分
完成死信队列配置和测试
开始做订单超时自动取消功能的技术方案设计
参加周四下午的分布式事务分享会
给新来的实习生做一次Git工作流培训

PR链接：
https://github.com/company/backend/pull/456 订单查询优化
https://github.com/company/backend/pull/461 头像上传修复
https://github.com/company/backend/pull/468 支付回调幂等
""",

    '公文': """\
市数据管理局要发一个关于推进政务数字化转型的通知

发给各区县人民政府，市政府各委、办、局

大概内容是这样的：

为了落实国务院关于数字政府建设的指导意见和省政府的工作部署，推动全市政务服务数字化转型，提升政府治理效能和公共服务水平，经市政府同意，现在要开展全市政务数字化转型工作。

工作目标有几个：
到2026年底，全市政务服务事项网上可办率要达到95%以上，全程网办率达到80%以上。建成统一的政务数据共享交换平台，实现市区两级数据互联互通。建设不少于50个典型数字化应用场景。群众办事平均跑动次数降到0.3次以下。

主要任务包括：

第一方面 政务服务平台升级改造
要升级市政务服务网和"城市通"App，实现PC端和移动端统一认证、统一入口。推行电子证照全面应用，实现身份证、营业执照、不动产权证等高频证照电子化。建设智能客服系统，提供7x24小时在线咨询服务。技术要求是前端用 Vue 3 + TypeScript，后端用 Spring Cloud 微服务架构，部署在市政务云平台上，接口规范遵循《政务服务平台接口规范》（GB/T 39554-2020）。

第二方面 数据共享交换平台建设
建设全市统一的数据目录，梳理政务数据资源不少于5000项。打通公安、民政、人社、住建、市场监管等重点部门的数据壁垒。建设数据质量监控系统，数据更新及时率达到98%以上。

第三方面 典型应用场景建设
重点打造"一件事一次办"主题集成服务，首批推出新生儿出生、企业开办、不动产交易、退休养老等20个主题。每个主题要实现材料减少50%以上，时限压缩60%以上。

实施步骤分三个阶段：
2026年5月到6月是动员部署阶段，成立工作专班、制定实施细则
2026年7月到11月是全面推进阶段，各项任务同步推进
2026年12月是总结验收阶段，组织考核评估

经费保障方面：
市财政安排专项经费预算如下：
政务平台升级改造 800万元
数据共享平台建设 500万元
应用场景开发 300万元
培训和运维 200万元
总计1800万元
各区县按照不低于市级标准1:0.5的比例配套资金

工作要求：
各区县各部门要高度重视，主要领导亲自抓。每月25号前报送工作进展。工作推进中遇到的重大问题及时报告。

联系人：张伟，电话：0571-88888888，邮箱：zhangwei@data.gov.example.cn

发文单位：市数据管理局
发文日期：2026年4月21日
文号大概是 数管发〔2026〕15号
抄送：市委办公厅，市人大常委会办公厅，市政协办公厅
""",
}


def load_config():
    if not os.path.isfile(CONFIG_PATH):
        print(f'错误：找不到配置文件 {CONFIG_PATH}', file=sys.stderr)
        sys.exit(1)
    with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
        cfg = json.load(f)
    if not (cfg.get('api_key') or os.getenv('MDTONG_AI_API_KEY') or os.getenv('ANTHROPIC_API_KEY')):
        print('错误：未配置 API Key，请在 config.json 中配置 api_key，或设置 MDTONG_AI_API_KEY/ANTHROPIC_API_KEY', file=sys.stderr)
        sys.exit(1)
    return cfg


def generate_one(doc_type, raw_text, api_key, model, max_retries=2):
    from converter.ai_polish import polish
    for attempt in range(max_retries + 1):
        try:
            return polish(
                raw_text=raw_text,
                api_key=api_key,
                model=model,
                template=doc_type,
                on_chunk=None,
            )
        except Exception as e:
            if attempt < max_retries:
                wait = 5 * (attempt + 1)
                print(f'    API 调用失败 (尝试 {attempt + 1}/{max_retries + 1}): {e}')
                print(f'    等待 {wait}s 后重试...')
                time.sleep(wait)
            else:
                raise


def save_sample(doc_type, content, output_dir=SAMPLES_DIR):
    os.makedirs(output_dir, exist_ok=True)
    filename = f'样本_{doc_type}.md'
    filepath = os.path.join(output_dir, filename)
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
    return filepath


def main():
    parser = argparse.ArgumentParser(description='为 MD通 生成 6 种文档类型的样本文档')
    parser.add_argument('--type', choices=list(SAMPLE_INPUTS.keys()),
                        help='仅生成指定类型的样本（默认生成全部）')
    parser.add_argument('--output-dir', default=SAMPLES_DIR,
                        help=f'输出目录（默认 {SAMPLES_DIR}）')
    args = parser.parse_args()

    cfg = load_config()
    api_key = cfg.get('api_key', '')
    model = cfg.get('ai_model', 'gpt-5.5')

    types_to_generate = [args.type] if args.type else list(SAMPLE_INPUTS.keys())

    print('=== MD通 样本文档生成器 ===')
    print(f'模型: {model}')
    print(f'待生成: {len(types_to_generate)} 种文档类型')
    print(f'输出目录: {args.output_dir}')
    print()

    results = []
    for i, doc_type in enumerate(types_to_generate, 1):
        print(f'[{i}/{len(types_to_generate)}] 正在生成「{doc_type}」样本...')
        start = time.time()
        try:
            raw_text = SAMPLE_INPUTS[doc_type] + FORMAT_HINT
            content, _usage = generate_one(doc_type, raw_text, api_key, model)
            filepath = save_sample(doc_type, content, args.output_dir)
            elapsed = time.time() - start
            print(f'    完成! 耗时 {elapsed:.1f}s, 保存至 {filepath}')
            results.append((doc_type, filepath, True, elapsed))
        except Exception as e:
            elapsed = time.time() - start
            print(f'    失败! 耗时 {elapsed:.1f}s, 错误: {e}')
            results.append((doc_type, '', False, elapsed))

    print()
    print('=== 生成结果汇总 ===')
    success = sum(1 for _, _, ok, _ in results if ok)
    print(f'成功: {success}/{len(results)}')
    for doc_type, filepath, ok, elapsed in results:
        status = 'OK' if ok else 'FAILED'
        print(f'  [{status}] {doc_type} ({elapsed:.1f}s) {filepath}')


if __name__ == '__main__':
    main()
