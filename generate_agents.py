#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成 12 个实验 Agent HTML 文件：6 品类 × 2 组别（协作/非协作规范诉求）。

分类说明（对照《AI对话截图.docx》商品外观与价格描述 + 实验要求）：
- 每个 Agent 对应一个「品类」，包含该品类的普通款与环保可降解款两类商品信息
  （外观描述 + 价格），例如「笔类」同时包含「金属中性笔」与「咖啡渣环保中性笔」；
- 协作组与非协作组仅在干预文案的关键词上区分
  （协作组：加入我们 / 一起 / 共同 / 我们；非协作组：请选择 / 您 / 个人）；
- 开场欢迎话术固定；
- 输入环保款商品 -> 提示输入普通版本商品（实验要求被试输入普通非环保商品）；
- 输入其他 / 不在实验商品范围内的输入 -> 固定报错文案。
"""

import os
import json

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(BASE_DIR, 'agents')

# ========== 品类数据（普通款 + 环保款 外观/价格，来自《商品外观与价格描述》） ==========
CATEGORIES = [
    {
        'category': '笔类',
        'copy_key': '金属中性笔',
        'normal_name': '金属中性笔',
        'normal_appearance': '不锈钢金属笔杆，盖帽式商务签字笔；有黑色、银色两种笔身，笔尾和笔夹位置带银色金属装饰，笔身简约修长',
        'normal_price': '5.68',
        'eco_name': '咖啡渣环保中性笔',
        'eco_appearance': '笔身由咖啡渣可降解材料制成，有白色、浅绿、深黑三色；搭配牛皮纸开窗简易礼盒包装，风格质朴简约，适合商务礼品',
        'eco_price': '10.00',
        'normal_aliases': ['金属中性笔', '中性笔', '金属笔', '签字笔'],
        'eco_aliases': ['咖啡渣环保中性笔', '咖啡渣外壳笔', '咖啡渣笔', '咖啡渣'],
    },
    {
        'category': '包装胶带',
        'copy_key': '透明胶带大卷',
        'normal_name': '得力透明封箱胶带',
        'normal_appearance': '透明淡黄色塑料大卷胶带，卷芯印有deli得力标识，胶带透光，适合快递打包封口',
        'normal_price': '7.11',
        'eco_name': '可降解封箱胶带',
        'eco_appearance': '浅棕米黄色，多卷捆绑；主打可降解材质，表面光滑，支持手撕',
        'eco_price': '11.00',
        'normal_aliases': ['透明胶带大卷', '得力透明封箱胶带', '透明胶带', '胶带大卷', '大卷胶带', '封箱胶带', '胶带'],
        'eco_aliases': ['可降解封箱胶带', '可降解胶带'],
    },
    {
        'category': '垃圾袋',
        'copy_key': '垃圾袋',
        'normal_name': '飞达三和普通垃圾袋',
        'normal_appearance': '黑色卷装塑料袋，多卷堆叠，加厚款，常规塑料垃圾袋',
        'normal_price': '11.20',
        'eco_name': '玉米淀粉全降解垃圾袋',
        'eco_appearance': '米白色背心式垃圾袋，多卷捆装；原料为玉米淀粉、PBAT+PLA全生物降解材质，袋身带标识',
        'eco_price': '14.75',
        'normal_aliases': ['垃圾袋', '飞达三和普通垃圾袋', '普通垃圾袋', '塑料袋', '袋子', '垃圾'],
        'eco_aliases': ['玉米淀粉全降解垃圾袋', '玉米淀粉垃圾袋', '玉米淀粉'],
    },
    {
        'category': '牙刷',
        'copy_key': '软毛牙刷',
        'normal_name': '普通软毛牙刷',
        'normal_appearance': '透明塑料手柄，有黑、棕、透粉三色；小刷头，黑色炭丝软毛，家用成人款',
        'normal_price': '12.90',
        'eco_name': '竹制环保牙刷',
        'eco_appearance': '天然竹木手柄，手柄印有品牌字样；搭配不同颜色竹炭软毛刷头，牛皮纸盒独立包装，共8支，风格天然质朴',
        'eco_price': '14.21',
        'normal_aliases': ['软毛牙刷', '普通软毛牙刷', '牙刷', '毛刷'],
        'eco_aliases': ['竹制环保牙刷', '竹牙刷', '竹制牙刷'],
    },
    {
        'category': '纸杯',
        'copy_key': '一次性纸杯',
        'normal_name': '永辉优选一次性纸杯',
        'normal_appearance': '白色袋装，袋身带橙色装饰线条与"Smile"字样，袋内装多摞白色一次性纸杯；容量255mL，加厚食品用纸，适用于家用、商务、会议',
        'normal_price': '20.90',
        'eco_name': '恒鑫PLA生物可降解纸杯',
        'eco_appearance': '纯白色双层中空纸杯，可搭配白色杯盖；采用PLA淋膜可降解材质，双层结构隔热',
        'eco_price': '25.90',
        'normal_aliases': ['一次性纸杯', '永辉优选一次性纸杯', '纸杯', '水杯', '杯子'],
        'eco_aliases': ['恒鑫PLA生物可降解纸杯', 'PLA可降解纸杯', 'PLA纸杯', '可降解纸杯', 'PLA', 'pla'],
    },
    {
        'category': '纸张',
        'copy_key': 'A4白纸',
        'normal_name': '超群A4白纸',
        'normal_appearance': '纯白色A4复印纸，整齐堆叠，纸面洁白无异味，可用于打印、复印、书写、绘画，也有A3规格可选',
        'normal_price': '5.80',
        'eco_name': '80克A4甘蔗纸',
        'eco_appearance': '浅米黄色A4纸，原料为甘蔗纤维，天然质朴色调；支持打印、书写、绘画，属于环保纸张',
        'eco_price': '7.51',
        'normal_aliases': ['A4白纸', '超群A4白纸', '白纸', 'A4纸', 'a4纸', '打印纸'],
        'eco_aliases': ['80克A4甘蔗纸', 'A4甘蔗纸', '甘蔗纸', '甘蔗'],
    },
]

# ========== 干预文案（与《实验参考文本》逐字一致，核心操纵句不可改动） ==========
COLLABORATION_COPY = {
    'A4白纸':   '加入我们，让我们一起为环保做出贡献！\n甘蔗纸采用环保甘蔗浆制成，减少林木砍伐、降低污水排放，无荧光增白剂，书写打印效果与普通纸张无异，兼顾办公实用与生态保护。已有65%的消费者选择了A4甘蔗纸。加入我们，让我们一起选择A4甘蔗纸，以每一次选择传递环保力量，让绿色办公成为共同常态。',
    '一次性纸杯': '加入我们，让我们一起为环保做出贡献！\nPLA可降解纸杯采用食品级可降解环保材料，无有害化学物质，耐高温、不易变形，日常饮水与办公使用安全便捷。生产环节减少塑料污染，助力资源循环，兼顾实用与环保。已有65%的消费者选择了PLA可降解纸杯。加入我们，让我们一起选择PLA可降解纸杯，以每一次选择传递环保力量，让绿色生活成为共同常态。',
    '金属中性笔': '加入我们，让我们一起为环保做出贡献！\n咖啡渣外壳笔采用可降解外壳与无毒墨水，书写流畅顺滑，适配日常学习、办公等多种场景，既满足使用需求，又共同减轻传统塑料笔带来的环境负担。已有65%的消费者选择了咖啡渣外壳笔。加入我们，让我们一起选择咖啡渣外壳笔，以每一次选择传递环保力量，让绿色办公成为共同常态。',
    '透明胶带大卷': '加入我们，让我们一起为环保做出贡献！\n可降解胶带采用环保可降解材质，材料可自然分解，不产生长期塑料污染，有效减轻环境负担。已有65%的消费者选择了可降解胶带。加入我们，让我们一起选择可降解胶带，以每一次选择传递环保力量，让绿色生活成为共同常态。',
    '垃圾袋':   '加入我们，让我们一起为环保做出贡献！\n玉米淀粉垃圾袋采用自然材料玉米淀粉制成，材料可自然分解，不产生长期塑料污染，有效减轻环境负担。已有65%的消费者选择了玉米淀粉垃圾袋。加入我们，让我们一起选择玉米淀粉垃圾袋，以每一次选择传递环保力量，让绿色成为生活共同常态。',
    '软毛牙刷': '加入我们，让我们一起为环保做出贡献！\n竹牙刷选用可降解竹制环保材质制成，材料可自然分解，不产生长期塑料污染，有效减轻环境负担。已有65%的消费者选择了竹牙刷。加入我们，让我们一起选择竹牙刷，以每一次选择传递环保力量，让绿色成为生活共同常态。'
}

NON_COLLABORATION_COPY = {
    'A4白纸':   '请选择绿色消费！\n甘蔗纸采用环保甘蔗浆制成，减少林木砍伐、降低污水排放，无荧光增白剂，书写打印效果与普通纸张无异，兼顾办公实用与生态保护。已有65%的消费者选择了A4甘蔗纸。请您选择A4甘蔗纸，以每一次选择传递环保力量，让绿色办公成为个人习惯。',
    '一次性纸杯': '请选择绿色消费！\nPLA可降解纸杯采用食品级可降解环保材料，无有害化学物质，耐高温、不易变形，日常饮水与办公使用安全便捷。生产环节减少塑料污染，助力资源循环，兼顾实用与环保。已有65%的消费者选择了PLA可降解纸杯。请您选择PLA可降解纸杯，以每一次选择传递环保力量，让绿色生活成为个人习惯。',
    '金属中性笔': '请选择绿色消费！\n咖啡渣外壳笔采用可降解外壳与无毒墨水，书写流畅顺滑，适配日常学习、办公等多种场景，既满足使用需求，又共同减轻传统塑料笔带来的环境负担。已有65%的消费者选择了咖啡渣外壳笔。请您选择咖啡渣外壳笔，以每一次选择传递环保力量，让绿色办公成为个人习惯。',
    '透明胶带大卷': '请选择绿色消费！\n可降解胶带采用环保可降解材质，材料可自然分解，不产生长期塑料污染，有效减轻环境负担。已有65%的消费者选择了可降解胶带。请您选择可降解胶带，以每一次选择传递环保力量，让绿色生活成为个人习惯。',
    '垃圾袋':   '请选择绿色消费！\n玉米淀粉垃圾袋采用自然材料玉米淀粉制成，材料可自然分解，不产生长期塑料污染，有效减轻环境负担。已有65%的消费者选择了玉米淀粉垃圾袋。请您选择玉米淀粉垃圾袋，以每一次选择传递环保力量，让绿色成为生活个人习惯。',
    '软毛牙刷': '请选择绿色消费！\n竹牙刷选用可降解竹制环保材质制成，材料可自然分解，不产生长期塑料污染，有效减轻环境负担。已有65%的消费者选择了竹牙刷。请您选择竹牙刷，以每一次选择传递环保力量，让绿色成为生活个人习惯。'
}

# ========== CSS（与原有样式一致，含 white-space:pre-line 以保留换行） ==========
CSS = """
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

:root {
  --green-dark: #2d5a3b;
  --green-mid: #4a8c5c;
  --green-light: #e8f5e9;
  --green-pale: #f1f8f2;
  --bg: #f5f7f5;
  --text-dark: #3a3a3a;
  --text-light: #ffffff;
  --bubble-user: #3d7a4f;
  --bubble-system: #e8ede8;
  --border: #d5ddd5;
  --shadow: 0 1px 3px rgba(0,0,0,0.08);
}

html, body {
  height: 100%;
  width: 100%;
  overflow: hidden;
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", "Helvetica Neue", sans-serif;
  font-size: 15px;
  color: var(--text-dark);
  background: var(--bg);
  -webkit-font-smoothing: antialiased;
}

.app-container {
  display: flex;
  flex-direction: column;
  height: 100%;
  max-width: 720px;
  margin: 0 auto;
  background: #ffffff;
  box-shadow: 0 0 20px rgba(0,0,0,0.04);
}

.app-header {
  flex-shrink: 0;
  padding: 14px 20px;
  background: linear-gradient(135deg, var(--green-dark), var(--green-mid));
  color: #fff;
  text-align: center;
  border-bottom: 2px solid #265035;
}
.app-header h1 {
  font-size: 17px;
  font-weight: 600;
  letter-spacing: 0.5px;
}

.chat-area {
  flex: 1;
  overflow-y: auto;
  padding: 16px 14px;
  display: flex;
  flex-direction: column;
  gap: 10px;
  scroll-behavior: smooth;
  -webkit-overflow-scrolling: touch;
}

.welcome-msg {
  align-self: flex-start;
  max-width: 82%;
  background: var(--bubble-system);
  color: var(--text-dark);
  padding: 10px 14px;
  border-radius: 16px 16px 16px 4px;
  font-size: 14px;
  line-height: 1.55;
  word-break: break-word;
  white-space: pre-line;
  box-shadow: var(--shadow);
}

.message {
  max-width: 82%;
  padding: 10px 14px;
  border-radius: 16px;
  font-size: 14px;
  line-height: 1.65;
  word-break: break-word;
  white-space: pre-line;
  animation: fadeIn 0.25s ease-out;
  box-shadow: var(--shadow);
}

.message.user {
  align-self: flex-end;
  background: var(--bubble-user);
  color: var(--text-light);
  border-bottom-right-radius: 4px;
}

.message.system {
  align-self: flex-start;
  background: var(--bubble-system);
  color: var(--text-dark);
  border-bottom-left-radius: 4px;
}

.message.loading {
  align-self: flex-start;
  background: var(--bubble-system);
  color: #888;
  font-style: italic;
  border-bottom-left-radius: 4px;
}

.message.error {
  align-self: flex-start;
  background: #fff3f0;
  color: #c0392b;
  border-bottom-left-radius: 4px;
  border: 1px solid #f5c6cb;
}

@keyframes fadeIn {
  from { opacity: 0; transform: translateY(6px); }
  to   { opacity: 1; transform: translateY(0); }
}

.input-area {
  flex-shrink: 0;
  display: flex;
  gap: 8px;
  padding: 10px 14px 14px;
  border-top: 1px solid var(--border);
  background: #fafbfa;
}

.input-area input {
  flex: 1;
  padding: 10px 14px;
  border: 1px solid var(--border);
  border-radius: 22px;
  font-size: 14px;
  outline: none;
  background: #fff;
  color: var(--text-dark);
  transition: border-color 0.2s;
}
.input-area input:focus { border-color: var(--green-mid); }
.input-area input::placeholder { color: #b5b5b5; }

.input-area button {
  flex-shrink: 0;
  padding: 0 22px;
  background: var(--green-mid);
  color: #fff;
  border: none;
  border-radius: 22px;
  font-size: 14px;
  font-weight: 500;
  cursor: pointer;
  transition: background 0.2s, opacity 0.2s;
  white-space: nowrap;
}
.input-area button:hover { background: var(--green-dark); }
.input-area button:active { opacity: 0.85; }
.input-area button:disabled { background: #a8c8a8; cursor: not-allowed; }

.chat-area::-webkit-scrollbar { width: 5px; }
.chat-area::-webkit-scrollbar-thumb { background: #c8d6c8; border-radius: 3px; }

@media (max-width: 480px) {
  .app-container { max-width: 100%; }
  .app-header { padding: 12px 16px; }
  .app-header h1 { font-size: 15px; }
  .chat-area { padding: 12px 10px; gap: 8px; }
  .message { max-width: 88%; padding: 9px 12px; font-size: 13px; }
  .input-area { padding: 8px 10px 10px; }
  .input-area input { padding: 8px 12px; font-size: 13px; }
  .input-area button { padding: 0 16px; font-size: 13px; }
}
"""

# ========== JS 逻辑（三分支：环保款提示 / 普通款商品信息+干预 / 其他报错） ==========
JS_TEMPLATE = """
(function() {
  'use strict';

  /* ========== Agent 身份 ========== */
  var AGENT_NAME = '小林';
  var MY_CATEGORY = __CATEGORY__;
  var MY_NORMAL_PRODUCT = __NORMAL_NAME__;
  var MY_ECO_PRODUCT = __ECO_NAME__;
  var MY_GROUP = __GROUP_LABEL__;
  var MY_NORMAL_ALIASES = __NORMAL_ALIASES__;
  var MY_ECO_ALIASES = __ECO_ALIASES__;
  var NORMAL_RESPONSE = __NORMAL_RESPONSE__;

  /* ========== 固定话术（实验要求，不可改动） ========== */
  var WELCOME_MSG = '顾客您好！欢迎光临小林超市，我是本店的 AI 智能助手小林，请输入您的商品信息。';
  var ERROR_MSG = '抱歉，暂不支持该商品，请输入本超市商品：A4 白纸、金属中性笔、透明胶带大卷、垃圾袋、软毛牙刷、一次性纸杯。';
  var ECO_PROMPT_MSG = '本次请输入普通版本商品，请重新输入对应普通商品名称。';

  /* ========== 全局状态 ========== */
  var sessionId = 'sess_' + Date.now().toString(36) + '_' + Math.random().toString(36).substr(2, 9);
  var chatArea = document.getElementById('chatArea');
  var userInput = document.getElementById('userInput');
  var sendBtn = document.getElementById('sendBtn');
  var interactionLog = [];

  /* ========== 初始化 ========== */
  function init() {
    showWelcome(WELCOME_MSG);
    reportToCredamo('session_init', { category: MY_CATEGORY, product: MY_NORMAL_PRODUCT, group: MY_GROUP });
    sendBtn.addEventListener('click', handleSend);
    userInput.addEventListener('keydown', function(e) {
      if (e.key === 'Enter') { e.preventDefault(); handleSend(); }
    });
    userInput.focus();
  }

  /* ========== 消息渲染 ========== */
  function showWelcome(text) {
    var div = document.createElement('div');
    div.className = 'welcome-msg';
    div.textContent = text;
    chatArea.appendChild(div);
    scrollToBottom();
  }

  function addUserMessage(text) {
    var div = document.createElement('div');
    div.className = 'message user';
    div.textContent = text;
    chatArea.appendChild(div);
    scrollToBottom();
  }

  function addLoadingMessage() {
    var div = document.createElement('div');
    div.className = 'message loading';
    div.textContent = AGENT_NAME + '正在为您查询…';
    div.id = 'loadingMsg';
    chatArea.appendChild(div);
    scrollToBottom();
    return div;
  }

  function replaceWithSystemMessage(loadingDiv, text) {
    var div = document.createElement('div');
    div.className = 'message system';
    div.textContent = text;
    loadingDiv.parentNode.replaceChild(div, loadingDiv);
    scrollToBottom();
    return div;
  }

  function scrollToBottom() {
    requestAnimationFrame(function() {
      chatArea.scrollTop = chatArea.scrollHeight;
    });
  }

  /* ========== 输入匹配逻辑 ========== */
  function matchAny(aliases, input) {
    var trimmed = input.trim();
    for (var i = 0; i < aliases.length; i++) {
      if (trimmed.indexOf(aliases[i]) !== -1) return true;
    }
    var lower = trimmed.toLowerCase();
    for (var j = 0; j < aliases.length; j++) {
      if (lower.indexOf(aliases[j].toLowerCase()) !== -1) return true;
    }
    return false;
  }

  /* ========== 数据回调 ========== */
  function reportToCredamo(eventType, data) {
    try {
      if (window.parent && window.parent !== window) {
        window.parent.postMessage({
          source: 'green_consumption_experiment',
          eventType: eventType,
          sessionId: sessionId,
          category: MY_CATEGORY,
          product: MY_NORMAL_PRODUCT,
          data: data,
          timestamp: new Date().toISOString()
        }, '*');
      }
    } catch (e) { }
    interactionLog.push({
      eventType: eventType, sessionId: sessionId,
      category: MY_CATEGORY, product: MY_NORMAL_PRODUCT, data: data,
      timestamp: new Date().toISOString()
    });
  }

  window.getExperimentData = function() {
    return { sessionId: sessionId, category: MY_CATEGORY, product: MY_NORMAL_PRODUCT, interactions: interactionLog };
  };

  /* ========== 发送处理 ========== */
  function handleSend() {
    var text = userInput.value.trim();
    if (!text) return;

    addUserMessage(text);
    reportToCredamo('user_input', { input: text });
    userInput.value = '';
    userInput.focus();

    var loadingDiv = addLoadingMessage();

    var reply;
    var replyType;

    // 1) 被试直接输入环保款商品 -> 提示改输普通版本
    if (matchAny(MY_ECO_ALIASES, text)) {
      reply = ECO_PROMPT_MSG;
      replyType = 'eco_product_prompt';
    }
    // 2) 输入本品类普通商品 -> 输出商品信息（普通款+环保款）+ 干预文案
    else if (matchAny(MY_NORMAL_ALIASES, text)) {
      reply = NORMAL_RESPONSE;
      replyType = 'product_response';
    }
    // 3) 其余输入 -> 固定报错
    else {
      reply = ERROR_MSG;
      replyType = 'error_response';
    }

    sendBtn.disabled = true;
    setTimeout(function() {
      replaceWithSystemMessage(loadingDiv, reply);
      reportToCredamo(replyType, { input: text });
      sendBtn.disabled = false;
    }, 300);
  }

  /* ========== 启动 ========== */
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
"""


def build_normal_response(cat, copy):
    """普通款输入 -> 展示普通款 + 环保款 外观/价格 + 干预文案（协作/非协作关键词区分）。"""
    return (
        '【%s】\n\n'
        'A（普通款）%s\n'
        '%s\n'
        '价格：%s元\n\n'
        'B（环保可降解款）%s\n'
        '%s\n'
        '价格：%s元\n\n'
        '%s' % (
            cat['category'],
            cat['normal_name'], cat['normal_appearance'], cat['normal_price'],
            cat['eco_name'], cat['eco_appearance'], cat['eco_price'],
            copy
        )
    )


def generate_html(group_type, cat):
    group_label = 'collaboration' if group_type == 'collaboration' else 'non_collaboration'
    copy_dict = COLLABORATION_COPY if group_type == 'collaboration' else NON_COLLABORATION_COPY
    copy = copy_dict[cat['copy_key']]
    normal_response = build_normal_response(cat, copy)

    js = (JS_TEMPLATE
          .replace('__CATEGORY__', json.dumps(cat['category'], ensure_ascii=False))
          .replace('__NORMAL_NAME__', json.dumps(cat['normal_name'], ensure_ascii=False))
          .replace('__ECO_NAME__', json.dumps(cat['eco_name'], ensure_ascii=False))
          .replace('__GROUP_LABEL__', json.dumps(group_label, ensure_ascii=False))
          .replace('__NORMAL_ALIASES__', json.dumps(cat['normal_aliases'], ensure_ascii=False))
          .replace('__ECO_ALIASES__', json.dumps(cat['eco_aliases'], ensure_ascii=False))
          .replace('__NORMAL_RESPONSE__', json.dumps(normal_response, ensure_ascii=False)))

    html = ('<!DOCTYPE html>\n'
            '<html lang="zh-CN">\n'
            '<head>\n'
            '<meta charset="UTF-8">\n'
            '<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">\n'
            '<title>绿色消费产品介绍</title>\n'
            '<style>\n' + CSS.strip() + '\n</style>\n'
            '</head>\n'
            '<body>\n\n'
            '<div class="app-container">\n'
            '  <header class="app-header">\n'
            '    <h1>绿色消费产品介绍</h1>\n'
            '  </header>\n\n'
            '  <div class="chat-area" id="chatArea"></div>\n\n'
            '  <div class="input-area">\n'
            '    <input type="text" id="userInput" placeholder="请输入产品名称…" autocomplete="off" enterkeyhint="send">\n'
            '    <button id="sendBtn">发送</button>\n'
            '  </div>\n'
            '</div>\n\n'
            '<script>\n' + js.strip() + '\n</script>\n'
            '</body>\n'
            '</html>\n')

    return html


def main():
    os.makedirs(OUT_DIR, exist_ok=True)

    groups = [
        ('collaboration', '协作组'),
        ('non_collaboration', '非协作组')
    ]

    count = 0
    for group_type, group_label in groups:
        for cat in CATEGORIES:
            count += 1
            filename = '%02d_%s_%s.html' % (count, group_label, cat['copy_key'])
            filepath = os.path.join(OUT_DIR, filename)

            html = generate_html(group_type, cat)
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(html)

            print('[%02d/12] %s  (%d bytes)' % (count, filename, len(html.encode('utf-8'))))

    print('\n全部 12 个 Agent HTML 文件已生成到: %s' % OUT_DIR)


if __name__ == '__main__':
    main()
