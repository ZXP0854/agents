#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成 12 个实验 Agent HTML 文件：6 品类 × 2 组别（协作/非协作规范诉求）。

分类说明（对照《AI对话截图.docx》商品外观与价格描述 + 实验要求）：
- 每个 Agent 对应一个「品类」，内置该品类的普通款与环保可降解款两类商品信息
  （外观描述 + 价格），例如「笔类」同时包含「金属中性笔」与「咖啡渣环保中性笔」；
- 文件名按品类命名：笔类 / 包装胶带 / 垃圾袋 / 牙刷 / 纸杯 / 纸张；
- 协作组与非协作组仅在干预文案的关键词上区分
  （协作组：加入我们 / 一起 / 共同 / 我们；非协作组：请选择 / 您 / 个人）；
- 开场欢迎话术固定；
- 输入普通商品 -> 「本店有两种X售卖：」+ 商品信息 + 干预文案；
- 输入环保款商品 -> 提示输入普通版本商品；
- 输入其他 -> 按品类固定报错（如「请输入与笔有关的商品」）；
- 被试追问（优点/好处/缺点/用途/价格/安全/耐用/对比/推荐等）-> 从 A/B 知识库回答，
  12 个 agent 回答风格保持一致，仅组别关键词不同。
"""

import os
import json

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(BASE_DIR, 'agents')

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

# ========== 品类数据（普通款 + 环保款 外观/价格/知识库，来自《商品外观与价格描述》） ==========
CATEGORIES = [
    {
        'category': '笔类',
        'measure': '笔',
        'copy_key': '金属中性笔',
        'normal_name': '金属中性笔',
        'normal_appearance': '不锈钢金属笔杆，盖帽式商务签字笔；有黑色、银色两种笔身，笔尾和笔夹位置带银色金属装饰，笔身简约修长',
        'normal_price': '5.68',
        'normal_aliases': ['金属中性笔', '中性笔', '金属笔', '签字笔', '笔'],
        'normal_kb': {
            '优点': '不锈钢金属笔杆，坚固耐用；商务外观简洁大方，笔身修长、手感沉稳；书写顺滑，适合日常办公与签字。',
            '好处': '价格实惠，一支即可长期使用；金属材质结实不易坏，性价比高；随处可买到，替换方便。',
            '缺点': '不锈钢笔杆冬天手感偏凉；重量比塑料笔重，长时间书写可能累手；外观颜色选择较少（仅黑色、银色）。',
            '用途': '日常办公书写、会议签字、笔记记录、商务场合使用。',
            '价格': '5.68元/支，价格实惠。',
            '安全': '不锈钢材质，不含塑化剂；正常书写安全，注意避免笔尖误伤。',
            '耐用': '金属笔杆坚固耐用、不易断裂；正常使用可用较长时间，笔芯可替换。',
        },
        'eco_name': '咖啡渣环保中性笔',
        'eco_appearance': '笔身由咖啡渣可降解材料制成，有白色、浅绿、深黑三色；搭配牛皮纸开窗简易礼盒包装，风格质朴简约，适合商务礼品',
        'eco_price': '10.00',
        'eco_aliases': ['咖啡渣环保中性笔', '咖啡渣外壳笔', '咖啡渣笔', '咖啡渣'],
        'eco_kb': {
            '优点': '笔身由废弃咖啡渣回收制作，实现废物利用；采用可降解材料，减少塑料污染；无毒墨水书写安全；自带咖啡自然色泽与纹理，外观独特有质感。',
            '好处': '使用环保材料，减轻环境负担；每支笔约利用5-10克废弃咖啡渣，让咖啡渣变废为宝；选择它即是在为减少塑料污染出一份力。',
            '缺点': '颜色以咖啡色系为主，选择较少；外壳硬度略低于ABS塑料笔，抗摔性稍弱；不同批次外观可能略有差异；市场渠道较少。',
            '用途': '日常学习办公书写、签字笔记；环保主题礼品、企业定制文具；适合日常书写。',
            '价格': '10.00元/支，比普通款贵一些，但环保价值更高。',
            '安全': '外壳经检测不含重金属与有害塑化剂；墨水符合国家文具安全标准，无毒无害；咖啡渣经过消毒处理。',
            '耐用': '正常书写寿命与普通中性笔相当；外壳在正常使用下不易断裂；使用完毕可更换笔芯，延长寿命。',
        },
        'compare': '金属中性笔更便宜、更结实；咖啡渣环保中性笔更环保、更有质感。两者书写体验基本一致，差别主要在外观材质与环保属性上。',
        'recommend': '如果您更看重环保与可持续，推荐选择B（咖啡渣环保中性笔）；如果更看重价格与耐用，可以选择A（金属中性笔）。',
    },
    {
        'category': '包装胶带',
        'measure': '胶带',
        'copy_key': '透明胶带大卷',
        'normal_name': '得力透明封箱胶带',
        'normal_appearance': '透明淡黄色塑料大卷胶带，卷芯印有deli得力标识，胶带透光，适合快递打包封口',
        'normal_price': '7.11',
        'normal_aliases': ['透明胶带大卷', '得力透明封箱胶带', '透明胶带', '胶带大卷', '大卷胶带', '封箱胶带', '胶带'],
        'normal_kb': {
            '优点': '透明度高、粘性强，封箱牢固；得力品牌，质量稳定；价格实惠，适合大量打包。',
            '好处': '便宜耐用，打包成本低；随处可买到；日常快递打包、封箱都很方便。',
            '缺点': '传统塑料材质，难以降解，废弃后会对环境造成长期负担。',
            '用途': '快递打包、纸箱封口、日常粘贴修补。',
            '价格': '7.11元/卷，价格实惠。',
            '安全': '常规胶带，正常使用安全；避免接触高温与明火。',
            '耐用': '粘性保持良好，常温下长期有效。',
        },
        'eco_name': '可降解封箱胶带',
        'eco_appearance': '浅棕米黄色，多卷捆绑；主打可降解材质，表面光滑，支持手撕',
        'eco_price': '11.00',
        'eco_aliases': ['可降解封箱胶带', '可降解胶带'],
        'eco_kb': {
            '优点': '采用可降解材质，使用后可自然分解，不产生长期塑料污染；粘性与传统胶带相当；表面光滑，支持手撕。',
            '好处': '减轻塑料污染，保护环境；降解后不残留微塑料；为绿色消费出一份力。',
            '缺点': '价格约为传统胶带的1.5-2倍；极端潮湿环境可能提前降解；大卷规格供应较少。',
            '用途': '快递包装、纸箱封口、手工制作、礼品包装、日常修补。',
            '价格': '11.00元/卷，比普通款贵一些。',
            '安全': '材料来自生物基材料，不含有毒物质；胶粘剂采用环保配方，VOC含量低。',
            '耐用': '常温下粘性保持与传统胶带相当；建议存放在干燥避光处。',
        },
        'compare': '得力透明封箱胶带更便宜、供应多；可降解封箱胶带更环保。两者粘性差别不大，主要差别在环保属性与价格上。',
        'recommend': '如果您更看重环保，推荐选择B（可降解封箱胶带）；如果更看重价格与供应便利，可以选择A（得力透明封箱胶带）。',
    },
    {
        'category': '垃圾袋',
        'measure': '垃圾袋',
        'copy_key': '垃圾袋',
        'normal_name': '飞达三和普通垃圾袋',
        'normal_appearance': '黑色卷装塑料袋，多卷堆叠，加厚款，常规塑料垃圾袋',
        'normal_price': '11.20',
        'normal_aliases': ['垃圾袋', '飞达三和普通垃圾袋', '普通垃圾袋', '塑料袋', '袋子', '垃圾'],
        'normal_kb': {
            '优点': '黑色加厚卷装，结实耐用；常规塑料材质，防水性好，装湿垃圾不易漏。',
            '好处': '加厚设计不易破；价格实惠，日常使用成本低。',
            '缺点': '传统塑料材质，降解需数百年，对环境负担大。',
            '用途': '家庭、办公室日常垃圾分类收集。',
            '价格': '11.20元/卷（多卷装）。',
            '安全': '常规塑料，正常使用安全；避免接触高温与明火。',
            '耐用': '加厚设计，承重好，不易破损。',
        },
        'eco_name': '玉米淀粉全降解垃圾袋',
        'eco_appearance': '米白色背心式垃圾袋，多卷捆装；原料为玉米淀粉、PBAT+PLA全生物降解材质，袋身带标识',
        'eco_price': '14.75',
        'eco_aliases': ['玉米淀粉全降解垃圾袋', '玉米淀粉垃圾袋', '玉米淀粉'],
        'eco_kb': {
            '优点': '原料为玉米淀粉等全生物降解材质，可自然分解；米白色背心式设计，使用方便；生产碳排放更低。',
            '好处': '降解后分解为水、二氧化碳和有机肥，不污染环境；摆脱石油资源依赖；减少白色污染。',
            '缺点': '成本约为普通款的2-3倍；遇水时间过长会软化，不适合长时间装湿垃圾；保质期约1-2年。',
            '用途': '家庭干垃圾分类收集、办公垃圾收集、户外活动临时垃圾收集（建议每日更换）。',
            '价格': '14.75元/卷（多卷装），比普通款贵。',
            '安全': '原料来自食用级玉米淀粉，不含有毒物质；不添加塑化剂。',
            '耐用': '干垃圾盛装表现良好；装湿垃圾建议不超过24小时。',
        },
        'compare': '飞达三和普通垃圾袋更便宜、更防水；玉米淀粉全降解垃圾袋更环保。两者承重接近，主要差别在降解性与价格上。',
        'recommend': '如果您更看重环保，推荐选择B（玉米淀粉全降解垃圾袋）；如果更看重价格与装湿垃圾，可以选择A（飞达三和普通垃圾袋）。',
    },
    {
        'category': '牙刷',
        'measure': '牙刷',
        'copy_key': '软毛牙刷',
        'normal_name': '普通软毛牙刷',
        'normal_appearance': '透明塑料手柄，有黑、棕、透粉三色；小刷头，黑色炭丝软毛，家用成人款',
        'normal_price': '12.90',
        'normal_aliases': ['软毛牙刷', '普通软毛牙刷', '牙刷', '毛刷'],
        'normal_kb': {
            '优点': '透明塑料手柄，三色可选；小刷头、炭丝软毛，清洁细腻；家用成人款，价格实惠。',
            '好处': '刷毛柔软，呵护牙龈；价格便宜，定期更换不心疼。',
            '缺点': '塑料手柄难以降解，废弃后对环境造成负担。',
            '用途': '日常早晚刷牙清洁。',
            '价格': '12.90元/支。',
            '安全': '食品级刷毛，符合口腔用品安全标准。',
            '耐用': '建议每3个月更换，正常使用期内刷毛不易变形。',
        },
        'eco_name': '竹制环保牙刷',
        'eco_appearance': '天然竹木手柄，手柄印有品牌字样；搭配不同颜色竹炭软毛刷头，牛皮纸盒独立包装，共8支，风格天然质朴',
        'eco_price': '14.21',
        'eco_aliases': ['竹制环保牙刷', '竹牙刷', '竹制牙刷'],
        'eco_kb': {
            '优点': '刷柄采用天然竹子，天然抗菌；竹炭软毛刷头，清洁护龈；牛皮纸盒独立包装。',
            '好处': '竹材生长快、可再生；废弃后刷柄可自然降解，减少塑料污染；天然材质更健康环保。',
            '缺点': '长期潮湿环境可能出现霉斑，需保持干燥；造型颜色选择较少；刷毛仍多为尼龙，需拆下分别处理。',
            '用途': '日常刷牙清洁、旅行携带、儿童环保教育、环保礼品。',
            '价格': '14.21元/支，比普通款略贵。',
            '安全': '竹材经消毒处理，不含有害物质；刷毛符合口腔用品安全标准。',
            '耐用': '保持干燥可正常使用3个月左右；避免长期浸泡。',
        },
        'compare': '普通软毛牙刷更便宜、颜色多；竹制环保牙刷更环保、天然抗菌。两者清洁效果接近，差别在材质与环保属性上。',
        'recommend': '如果您更看重环保，推荐选择B（竹制环保牙刷）；如果更看重价格与颜色，可以选择A（普通软毛牙刷）。',
    },
    {
        'category': '纸杯',
        'measure': '纸杯',
        'copy_key': '一次性纸杯',
        'normal_name': '永辉优选一次性纸杯',
        'normal_appearance': '白色袋装，袋身带橙色装饰线条与"Smile"字样，袋内装多摞白色一次性纸杯；容量255mL，加厚食品用纸，适用于家用、商务、会议',
        'normal_price': '20.90',
        'normal_aliases': ['一次性纸杯', '永辉优选一次性纸杯', '纸杯', '水杯', '杯子'],
        'normal_kb': {
            '优点': '255mL容量，加厚食品用纸；白色袋装，干净卫生；适用于家用、商务、会议。',
            '好处': '加厚不易烫手；价格实惠，聚会会议使用方便。',
            '缺点': '内壁PE淋膜，降解慢，对环境负担较大。',
            '用途': '家庭、办公室、会议、聚会饮水。',
            '价格': '20.90元/袋（多只装）。',
            '安全': '食品用纸，符合食品接触材料安全标准。',
            '耐用': '加厚设计，装热水不易软化变形。',
        },
        'eco_name': '恒鑫PLA生物可降解纸杯',
        'eco_appearance': '纯白色双层中空纸杯，可搭配白色杯盖；采用PLA淋膜可降解材质，双层结构隔热',
        'eco_price': '25.90',
        'eco_aliases': ['恒鑫PLA生物可降解纸杯', 'PLA可降解纸杯', 'PLA纸杯', '可降解纸杯', 'PLA', 'pla'],
        'eco_kb': {
            '优点': '纯白色双层中空设计，隔热不烫手；PLA淋膜可降解材质；可搭配杯盖。',
            '好处': 'PLA材料可降解，减少塑料污染；双层结构隔热，使用更舒适；更环保健康。',
            '缺点': '价格比普通款贵；需在工业堆肥条件下才能快速降解；耐油性略弱。',
            '用途': '办公室、会议、咖啡奶茶外带、注重环保的活动现场。',
            '价格': '25.90元/袋（多只装）。',
            '安全': '原料为植物基PLA，不含双酚A等有害物质；符合食品接触材料安全标准。',
            '耐用': '双层结构隔热，装热水1-2小时内不易软化；不宜微波炉加热。',
        },
        'compare': '永辉优选一次性纸杯更便宜；恒鑫PLA生物可降解纸杯更环保、双层隔热。两者容量相近，差别在材质与环保属性上。',
        'recommend': '如果您更看重环保与隔热，推荐选择B（恒鑫PLA生物可降解纸杯）；如果更看重价格，可以选择A（永辉优选一次性纸杯）。',
    },
    {
        'category': '纸张',
        'measure': '纸',
        'copy_key': 'A4白纸',
        'normal_name': '超群A4白纸',
        'normal_appearance': '纯白色A4复印纸，整齐堆叠，纸面洁白无异味，可用于打印、复印、书写、绘画，也有A3规格可选',
        'normal_price': '5.80',
        'normal_aliases': ['A4白纸', '超群A4白纸', '白纸', 'A4纸', 'a4纸', '打印纸', '纸张', '纸'],
        'normal_kb': {
            '优点': '纯白色A4复印纸，洁白无异味；可用于打印、复印、书写、绘画；有A3规格可选。',
            '好处': '白度高、打印效果清晰；价格实惠，办公学习常用。',
            '缺点': '传统木浆纸，依赖木材；漂白生产有一定环境负担。',
            '用途': '打印、复印、书写、绘画。',
            '价格': '5.80元/包。',
            '安全': '无异味，正常使用安全。',
            '耐用': '干燥环境下可长期保存。',
        },
        'eco_name': '80克A4甘蔗纸',
        'eco_appearance': '浅米黄色A4纸，原料为甘蔗纤维，天然质朴色调；支持打印、书写、绘画，属于环保纸张',
        'eco_price': '7.51',
        'eco_aliases': ['80克A4甘蔗纸', 'A4甘蔗纸', '甘蔗纸', '甘蔗'],
        'eco_kb': {
            '优点': '原料为甘蔗纤维，天然质朴色调；支持打印、书写、绘画；减少林木砍伐。',
            '好处': '利用甘蔗渣变废为宝；生产碳排放更低、无荧光增白剂；选择它即是在保护森林资源。',
            '缺点': '颜色略偏米黄，不如漂白纸白；价格略高；供应渠道不如普通纸广泛。',
            '用途': '打印、复印、书写、绘画、环保主题宣传材料。',
            '价格': '7.51元/包。',
            '安全': '不含荧光增白剂，对健康更友好。',
            '耐用': '干燥环境下可长期保存，书写打印效果与普通纸相当。',
        },
        'compare': '超群A4白纸更白、更便宜；80克A4甘蔗纸更环保、更健康。两者书写打印效果基本一致，差别在颜色与环保属性上。',
        'recommend': '如果您更看重环保与健康，推荐选择B（80克A4甘蔗纸）；如果更看重白度与价格，可以选择A（超群A4白纸）。',
    },
]

# ========== 环保概念科普词库（被试询问概念时的统一回答） ==========
ECO_CONCEPTS = {
    '可降解': '可降解材料是指在自然环境中能被微生物分解的材料，不会像传统塑料那样存留数百年。常见的可降解材料包括PLA（聚乳酸）、玉米淀粉、甘蔗浆、竹纤维等。',
    'PLA': 'PLA（聚乳酸）是由玉米、木薯等可再生植物资源发酵制成的生物降解塑料，生产过程的温室气体排放比传统塑料低约60%，在工业堆肥条件下可完全降解。',
    '咖啡渣': '全球每年产生约600万吨废弃咖啡渣。回收咖啡渣制作产品外壳，既减少了废弃物，又赋予了产品独特的质感与环保属性，是废物循环利用的典型案例。',
    '甘蔗浆': '甘蔗榨糖后剩余的纤维残渣（蔗渣）是制作甘蔗纸的主要原料。传统上蔗渣常被焚烧处理，将其转化为纸制品既实现了废物利用，又能减少对木材的依赖、保护森林资源。',
    '竹纤维': '竹子生长速度极快、可再生性强，竹纤维制品具有天然抗菌、可生物降解等特性，是塑料制品的理想替代材料，降解时间从数百年缩短至数年。',
    '玉米淀粉': '玉米淀粉是一种可再生资源，通过发酵和聚合工艺可制成生物塑料。它在工业堆肥条件下可完全降解，降解产物为水、二氧化碳和有机肥料，不会造成二次污染。',
    '塑料污染': '传统塑料制品在自然环境中需要200-500年才能完全降解，每年约有800万吨塑料流入海洋。选择可降解替代产品是减少塑料污染的有效方式之一。',
    '碳足迹': '碳足迹是指个人、组织或产品在整个生命周期中直接或间接产生的温室气体排放总量。选择环保材料制成的产品通常能显著降低碳足迹。',
    '循环经济': '循环经济是一种以资源高效利用和循环利用为核心的经济模式，强调"减量化、再利用、资源化"，追求产品使用结束后能够重新进入生产循环。',
    '绿色消费': '绿色消费是一种可持续的消费方式，强调在购买和使用产品时关注环保、节能和健康，通过每一次消费选择为减少碳排放、保护自然资源做出贡献。',
}

ECO_CONCEPT_KEYWORDS = {
    '可降解': '可降解', '降解': '可降解',
    'PLA': 'PLA', 'pla': 'PLA', '聚乳酸': 'PLA',
    '咖啡渣': '咖啡渣',
    '甘蔗': '甘蔗浆', '甘蔗浆': '甘蔗浆', '蔗渣': '甘蔗浆',
    '竹纤维': '竹纤维', '竹': '竹纤维',
    '玉米淀粉': '玉米淀粉', '玉米': '玉米淀粉',
    '塑料污染': '塑料污染', '塑料': '塑料污染', '污染': '塑料污染',
    '碳足迹': '碳足迹', '碳排放': '碳足迹',
    '循环经济': '循环经济', '循环': '循环经济',
    '绿色消费': '绿色消费', '绿色': '绿色消费',
    '环保': '绿色消费', '环保吗': '绿色消费', '环保性': '绿色消费', '环保在哪': '绿色消费',
}

# ========== 扩展知识库（被试可能追问的更多维度，每品类 A/B 各 10 项） ==========
EXTRA_KB = {
    '笔类': {
        'normal': {
            '材质': '不锈钢金属笔杆，金属材质。',
            '规格': '标准中性笔长度，笔杆修长，盖帽式设计。',
            '效果': '书写顺滑，出墨均匀，日常办公签字效果良好。',
            '性价比': '5.68元/支，价格实惠，性价比高。',
            '口碑': '商务办公常见选择，销量稳定，口碑良好。',
            '环保意义': '金属材质虽可回收，但生产过程能耗较高，环保意义有限。',
            '降解': '不锈钢材质难以自然降解，废弃后需回收处理。',
            '购买渠道': '文具店、超市、电商平台均有销售，购买方便。',
            '适用人群': '适合学生、上班族、商务人士日常书写。',
            '注意事项': '笔尖较细，注意避免摔落损坏笔尖。',
        },
        'eco': {
            '材质': '笔身由回收咖啡渣与可降解树脂制成。',
            '规格': '标准中性笔长度，白色、浅绿、深黑三色，牛皮纸礼盒包装。',
            '效果': '书写流畅度与传统中性笔相当，顺滑不卡顿。',
            '性价比': '10.00元/支，略贵但兼顾环保价值。',
            '口碑': '环保文具新选择，受到环保意识较强的消费者欢迎。',
            '环保意义': '变废为宝利用咖啡渣，减少塑料使用与环境污染。',
            '降解': '外壳可降解，降解时间远短于传统塑料笔。',
            '购买渠道': '电商平台、环保文具店有售，渠道正在拓展。',
            '适用人群': '适合注重环保的学生、上班族及商务礼品需求。',
            '注意事项': '避免长期泡水，以免外壳软化。',
        },
    },
    '包装胶带': {
        'normal': {
            '材质': '传统BOPP塑料材质。',
            '规格': '大卷装，透明淡黄色，卷芯印有得力标识。',
            '效果': '粘性强、透光，封箱牢固。',
            '性价比': '7.11元/卷，价格实惠，打包成本低。',
            '口碑': '得力品牌，质量稳定，广受好评。',
            '环保意义': '传统塑料难降解，环保意义较弱。',
            '降解': '塑料材质需数百年降解，对环境负担较大。',
            '购买渠道': '文具店、超市、电商平台随处可买。',
            '适用人群': '适合快递打包、办公封箱需求。',
            '注意事项': '避免高温与明火，注意储存防潮。',
        },
        'eco': {
            '材质': '植物基可降解材质，表面光滑。',
            '规格': '大卷装，浅棕米黄色，支持手撕。',
            '效果': '粘性与传统胶带相当，满足日常包装需求。',
            '性价比': '11.00元/卷，略贵但环保价值高。',
            '口碑': '环保包装新选择，受到绿色物流欢迎。',
            '环保意义': '可自然分解，减少塑料污染。',
            '降解': '堆肥条件下6-12个月可降解，不残留微塑料。',
            '购买渠道': '电商平台、环保用品店有售。',
            '适用人群': '适合注重环保的快递、电商与家庭用户。',
            '注意事项': '避免极端潮湿环境，以免提前降解。',
        },
    },
    '垃圾袋': {
        'normal': {
            '材质': '传统PE塑料，黑色加厚。',
            '规格': '卷装多卷堆叠，加厚款。',
            '效果': '结实防水，装湿垃圾不易漏。',
            '性价比': '11.20元，价格实惠。',
            '口碑': '常见品牌，销量稳定。',
            '环保意义': '塑料难降解，环保意义较弱。',
            '降解': '需数百年降解。',
            '购买渠道': '超市、电商平台随处可买。',
            '适用人群': '家庭、办公室日常垃圾收集。',
            '注意事项': '避免接触高温与明火。',
        },
        'eco': {
            '材质': '玉米淀粉、PBAT、PLA全生物降解材质。',
            '规格': '米白色背心式，多卷捆装，袋身带标识。',
            '效果': '韧性与承重能力接近传统垃圾袋。',
            '性价比': '14.75元，略贵但降解快。',
            '口碑': '环保垃圾袋新选择，受欢迎。',
            '环保意义': '降解产物为水、二氧化碳和有机肥，不污染环境。',
            '降解': '堆肥条件下3-6个月可完全降解。',
            '购买渠道': '电商平台、环保超市有售。',
            '适用人群': '注重环保的家庭、办公用户。',
            '注意事项': '不宜长时间装湿垃圾，建议每日更换。',
        },
    },
    '牙刷': {
        'normal': {
            '材质': '透明塑料手柄，炭丝软毛刷头。',
            '规格': '小刷头，黑、棕、透粉三色，家用成人款。',
            '效果': '软毛清洁细腻，呵护牙龈。',
            '性价比': '12.90元/支，价格实惠。',
            '口碑': '常见家用牙刷，销量稳定。',
            '环保意义': '塑料手柄难降解，环保意义较弱。',
            '降解': '塑料手柄需数百年降解。',
            '购买渠道': '超市、便利店、电商平台随处可买。',
            '适用人群': '成人日常口腔清洁。',
            '注意事项': '建议每3个月更换。',
        },
        'eco': {
            '材质': '天然竹木手柄，竹炭软毛刷头。',
            '规格': '成人款，牛皮纸盒独立包装，共8支。',
            '效果': '清洁效果与普通牙刷相当，竹炭软毛护龈。',
            '性价比': '14.21元/支，略贵但环保。',
            '口碑': '环保牙刷热门选择。',
            '环保意义': '竹材可再生、可降解，减少塑料污染。',
            '降解': '刷柄数年内自然降解，刷毛需拆下分别处理。',
            '购买渠道': '电商平台、环保生活店有售。',
            '适用人群': '注重环保的人群及儿童环保教育。',
            '注意事项': '保持干燥通风，避免长期浸泡发霉。',
        },
    },
    '纸杯': {
        'normal': {
            '材质': '加厚食品用纸，内壁PE淋膜。',
            '规格': '255mL容量，白色袋装多只。',
            '效果': '加厚防烫，盛装热水方便。',
            '性价比': '20.90元/袋，价格实惠。',
            '口碑': '永辉优选，常见家用品牌。',
            '环保意义': 'PE淋膜难降解，环保意义较弱。',
            '降解': '内壁PE淋膜降解较慢。',
            '购买渠道': '超市、电商平台随处可买。',
            '适用人群': '家庭、商务、会议使用。',
            '注意事项': '避免长时间盛装过烫液体。',
        },
        'eco': {
            '材质': 'PLA淋膜可降解材质，纯白双层中空。',
            '规格': '容量适中，可搭配白色杯盖。',
            '效果': '双层隔热不烫手，使用舒适。',
            '性价比': '25.90元/袋，略贵但环保。',
            '口碑': '环保纸杯热门选择。',
            '环保意义': 'PLA可降解，减少塑料污染。',
            '降解': '工业堆肥条件下可降解。',
            '购买渠道': '电商平台、环保用品店有售。',
            '适用人群': '注重环保的办公、聚会、外带需求。',
            '注意事项': '不宜微波炉加热。',
        },
    },
    '纸张': {
        'normal': {
            '材质': '传统木浆纸，纯白色。',
            '规格': 'A4规格，整齐堆叠，有A3规格可选。',
            '效果': '洁白无异味，打印复印清晰。',
            '性价比': '5.80元/包，价格实惠。',
            '口碑': '常见办公用纸，销量稳定。',
            '环保意义': '依赖木材砍伐，环保意义较弱。',
            '降解': '纸张本身可降解，但漂白生产有一定环境负担。',
            '购买渠道': '文具店、超市、电商平台随处可买。',
            '适用人群': '打印、复印、书写、绘画需求。',
            '注意事项': '防潮保存。',
        },
        'eco': {
            '材质': '甘蔗纤维，浅米黄色。',
            '规格': '80克A4规格。',
            '效果': '书写打印效果与普通纸无异。',
            '性价比': '7.51元/包，略贵但环保。',
            '口碑': '环保纸张热门选择。',
            '环保意义': '甘蔗渣变废为宝，减少林木砍伐，无荧光增白剂。',
            '降解': '可生物降解，对环境友好。',
            '购买渠道': '电商平台、环保文具店有售。',
            '适用人群': '打印、书写、绘画及环保宣传材料。',
            '注意事项': '防潮保存，避免长期暴晒。',
        },
    },
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

# ========== JS 逻辑 ==========
JS_TEMPLATE = """
(function() {
  'use strict';

  /* ========== Agent 身份 ========== */
  var AGENT_NAME = '小林';
  var MY_CATEGORY = __CATEGORY__;
  var MY_MEASURE = __MEASURE__;
  var MY_NORMAL_PRODUCT = __NORMAL_NAME__;
  var MY_ECO_PRODUCT = __ECO_NAME__;
  var MY_GROUP = __GROUP_LABEL__;
  var MY_NORMAL_ALIASES = __NORMAL_ALIASES__;
  var MY_ECO_ALIASES = __ECO_ALIASES__;
  var MY_COPY = __COPY__;
  var PRODUCT_INFO = __PRODUCT_INFO__;

  /* ========== A/B 知识库 ========== */
  var KB_A = __KB_A__;
  var KB_B = __KB_B__;
  var KB_COMPARE = __KB_COMPARE__;
  var KB_RECOMMEND = __KB_RECOMMEND__;

  /* ========== 追问收尾（核心操纵元素自然融入，避免生硬重复整句） ========== */
  var CLOSING_BY_TYPE = {
    collaboration: {
      '优点': '它（们）的这些优点，正是我们共同选择环保产品、为环保做出贡献的理由。',
      '好处': '这也让我们在为环保做出贡献的同时，收获更多。',
      '缺点': '虽然有一些不足，但为环保计，它仍值得我们一起选择。',
      '用途': '日常使用中，我们也能一起为环保尽一份力。',
      '价格': '多付一点，是我们一起为环保尽的心意。',
      '安全': '健康又环保，值得我们一起选择。',
      '耐用': '健康又环保，值得我们一起选择。',
      '材质': '选择这样的环保材质，正是我们为环保做出的选择。',
      '规格': '日常使用中，我们也能一起为环保尽一份力。',
      '效果': '这些效果，正是我们共同选择环保产品的理由。',
      '性价比': '多付一点，是我们一起为环保尽的心意。',
      '口碑': '越来越多的人加入我们，一起选择环保产品。',
      '环保意义': '这正是我们一起为环保做出贡献的意义。',
      '降解': '选择可降解材质，正是我们一起为环保尽的一份力。',
      '购买渠道': '欢迎加入我们，一起选择环保款。',
      '适用人群': '无论哪类人群，都能一起为环保做出贡献。',
      '注意事项': '虽然有一些注意点，但为环保计，它仍值得我们一起选择。',
      '对比': '综合来看，加入我们一起选择环保款，就是为环保做出贡献。',
      '推荐': '加入我们，一起为环保做出贡献。',
      'eco': '这也正是我们共同守护环境的方式。'
    },
    non_collaboration: {
      '优点': '这些优点，正是绿色消费的优选理由。',
      '好处': '这也是您为绿色消费尽的一份力。',
      '缺点': '虽然有一些不足，但为环保计，您仍可考虑它。',
      '用途': '日常使用中，您也能为绿色消费尽一份力。',
      '价格': '多付一点，是您为环保尽的心意。',
      '安全': '健康又环保，值得您选择。',
      '耐用': '健康又环保，值得您选择。',
      '材质': '选择这样的环保材质，是您为环保做出的选择。',
      '规格': '日常使用中，您也能为绿色消费尽一份力。',
      '效果': '这些效果，正是绿色消费的优选理由。',
      '性价比': '多付一点，是您为环保尽的心意。',
      '口碑': '越来越多的人选择绿色消费。',
      '环保意义': '这正是绿色消费的意义所在。',
      '降解': '选择可降解材质，是您为环保尽的一份力。',
      '购买渠道': '欢迎您选择环保款。',
      '适用人群': '无论哪类人群，都能为绿色消费尽一份力。',
      '注意事项': '虽然有一些注意点，但为环保计，您仍可考虑它。',
      '对比': '综合来看，请您选择绿色消费，环保款更值得考虑。',
      '推荐': '请您选择绿色消费，环保款值得考虑。',
      'eco': '这也正是您践行绿色消费的方式。'
    }
  };

  function getClosing(type) {
    var map = CLOSING_BY_TYPE[MY_GROUP === 'collaboration' ? 'collaboration' : 'non_collaboration'];
    return map[type] || map['推荐'];
  }

  /* ========== 固定话术（实验要求，不可改动） ========== */
  var WELCOME_MSG = '顾客您好！欢迎光临小林超市，我是本店的 AI 智能助手小林，请输入您的商品信息。';
  var ERROR_MSG = '抱歉，暂不支持该商品，请输入与' + MY_MEASURE + '有关的商品。';
  var ECO_PROMPT_MSG = '本次请输入普通版本商品，请重新输入对应普通商品名称。';
  var FALLBACK_QA_MSG = '关于这个问题我暂时没有更多信息。您可以向我询问这两款' + MY_MEASURE + '的：优点、缺点、好处、用途、价格、材质、规格、效果、性价比、口碑、环保意义、降解情况、购买渠道、适用人群、注意事项等。';
  var HELP_MSG = '您可以向我询问这两款' + MY_MEASURE + '的以下具体信息：\\n1. 优点、好处、缺点；\\n2. 用途、价格、材质、规格、效果；\\n3. 性价比、口碑、安全性、耐用性；\\n4. 环保意义、降解情况、购买渠道、适用人群、注意事项；\\n5. 两款商品的对比，或让我为您推荐。\\n您也可以了解环保知识，例如"PLA是什么""可降解是什么意思"等。';

  /* ========== 输入框灰字（提示语，随流程三阶段切换） ========== */
  var PLACEHOLDER_INITIAL = '请输入产品名称';
  var PLACEHOLDER_ASK_PRODUCT = '询问本店可购买的商品';
  var PLACEHOLDER_ACTIVE = '您现在可以询问两种商品的具体信息';

  function setPlaceholder(state) {
    if (state === 1) {
      userInput.placeholder = PLACEHOLDER_ASK_PRODUCT;
    } else if (state >= 2) {
      userInput.placeholder = PLACEHOLDER_ACTIVE;
    } else {
      userInput.placeholder = PLACEHOLDER_INITIAL;
    }
  }

  /* ========== 询问商品意图检测（阶段2） ========== */
  var ASK_PRODUCT_KEYS = ['有什么商品', '有哪些商品', '有什么', '有哪些', '卖什么', '卖哪些', '卖的东西', '售卖', '可购买', '本店有什么', '本店有哪些', '商品有哪些', '商品有什么', '商品', '卖'];

  function isAskProduct(input) {
    var lower = input.trim().toLowerCase();
    for (var i = 0; i < ASK_PRODUCT_KEYS.length; i++) {
      if (lower.indexOf(ASK_PRODUCT_KEYS[i].toLowerCase()) !== -1) return true;
    }
    return false;
  }

  /* ========== 追问类型（12 个 agent 风格统一，具体类型优先） ========== */
  var QUESTION_TYPES = [
    { type: '帮助', target: false, keys: ['能问什么', '可以问什么', '能问啥', '可以问啥', '问什么', '问些什么', '问点啥', '问哪些', '能问哪些', '可以问哪些', '问哪些问题', '能问的问题', '可以问的问题', '怎么问', '能问吗', '可以问吗', '能咨询什么', '可以咨询什么', '请问能问', '能了解什么', '可以了解什么'] },
    { type: '推荐', target: false, keys: ['推荐', '建议', '买哪个', '选哪个', '应该买', '值得买', '买什么', '怎么选', '买哪款', '选哪款'] },
    { type: '对比', target: false, keys: ['对比', '区别', '比较', '相比', 'vs', '哪个好', '有什么不同', '差别', '不一样', '哪个更好', '不同'] },
    { type: '价格', target: true,  keys: ['价格', '多少钱', '贵', '便宜', '成本', '花费', '划算', '价位'] },
    { type: '材质', target: true,  keys: ['材质', '材料', '成分', '什么做的', '原料'] },
    { type: '规格', target: true,  keys: ['规格', '尺寸', '容量', '多大', '多少毫升', '多少张', '多少只', '型号'] },
    { type: '效果', target: true,  keys: ['效果', '好不好用', '好用吗', '管用', '好用'] },
    { type: '性价比', target: true,  keys: ['性价比', '值不值', '值吗', '值这个价', '划得来'] },
    { type: '口碑', target: true,  keys: ['口碑', '评价', '销量', '受欢迎', '卖得好', '人气', '好评'] },
    { type: '环保意义', target: true,  keys: ['环保', '为什么环保', '环保在哪', '环保吗', '绿色', '对环境', '环境友好', '环保性'] },
    { type: '降解', target: true,  keys: ['降解', '分解', '多久', '降解时间', '分解时间', '能分解', '可降解'] },
    { type: '购买渠道', target: true,  keys: ['哪里买', '购买渠道', '怎么买', '去哪买', '在哪买', '哪里卖', '渠道'] },
    { type: '适用人群', target: true,  keys: ['适合谁', '适用人群', '什么人', '谁用', '人群', '适合什么人', '小孩', '儿童'] },
    { type: '注意事项', target: true,  keys: ['注意', '注意事项', '注意什么', '注意点', '使用注意', '提醒', '储存', '保存', '存放', '保质'] },
    { type: '优点', target: true,  keys: ['优点', '优势', '特色', '亮点', '哪里好', '好在', '好不好'] },
    { type: '好处', target: true,  keys: ['好处', '有什么用', '有什么好', '作用', '帮助'] },
    { type: '缺点', target: true,  keys: ['缺点', '不足', '劣势', '缺陷', '坏处', '不好', '问题'] },
    { type: '用途', target: true,  keys: ['用途', '怎么用', '使用场景', '能做什么', '适用', '适合', '场景', '干什么用', '干嘛用'] },
    { type: '安全', target: true,  keys: ['安全', '有毒', '无害', '健康', '放心', '标准', '检测', '危害'] },
    { type: '耐用', target: true,  keys: ['耐用', '质量', '寿命', '持久', '容易坏', '结实', '能用多久', '能用多长时间'] }
  ];

  var LABEL_MAP = { '优点': '优点', '好处': '好处', '缺点': '缺点', '用途': '用途', '价格': '价格', '安全': '安全性', '耐用': '耐用性', '材质': '材质', '规格': '规格', '效果': '效果', '性价比': '性价比', '口碑': '口碑', '环保意义': '环保意义', '降解': '降解情况', '购买渠道': '购买渠道', '适用人群': '适用人群', '注意事项': '注意事项' };

  /* ========== 环保概念科普词库 ========== */
  var ECO_CONCEPTS = __ECO_CONCEPTS__;
  var ECO_CONCEPT_KEYWORDS = __ECO_CONCEPT_KEYWORDS__;

  var A_TARGET_MARKERS = MY_NORMAL_ALIASES.concat(['非环保', '普通款', '普通', 'a款', '选a', '买a', 'a的', 'a有', 'a好', 'a是', 'a什么']);
  var B_TARGET_MARKERS = MY_ECO_ALIASES.concat(['环保产品', '环保款', '可降解款', '绿色款', '环保', '可降解', 'b款', '选b', '买b', 'b的', 'b有', 'b好', 'b是', 'b什么']);

  /* ========== 全局状态 ========== */
  var sessionId = 'sess_' + Date.now().toString(36) + '_' + Math.random().toString(36).substr(2, 9);
  var chatArea = document.getElementById('chatArea');
  var userInput = document.getElementById('userInput');
  var sendBtn = document.getElementById('sendBtn');
  var interactionLog = [];
  // 流程阶段：0=待输入商品；1=已展示干预文案，待询问商品信息；2=已展示商品信息，进入自由追问
  var flowStage = 0;

  /* ========== 初始化 ========== */
  function init() {
    showWelcome(WELCOME_MSG);
    setPlaceholder(0);
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

  /* ========== 追问检测 ========== */
  function detectQuestion(input) {
    var trimmed = input.trim();
    // 纯商品名（完整名）输入不作为追问，交由商品识别分支处理
    if (trimmed === MY_NORMAL_PRODUCT || trimmed === MY_ECO_PRODUCT) return null;
    var lower = trimmed.toLowerCase();

    var conceptKeys = ['PLA', '聚乳酸', '咖啡渣', '甘蔗浆', '蔗渣', '竹纤维', '玉米淀粉', '塑料污染', '碳足迹', '循环经济', '绿色消费'];

    // 1) 追问类型匹配（具体类型优先：优点/价格/材质/环保意义…）
    for (var i = 0; i < QUESTION_TYPES.length; i++) {
      var qt = QUESTION_TYPES[i];
      for (var j = 0; j < qt.keys.length; j++) {
        if (lower.indexOf(qt.keys[j].toLowerCase()) !== -1) return qt;
      }
    }

    // 2) 含明确疑问词的环保概念科普（如「PLA是什么」「咖啡渣是什么」）
    var askWords = ['是什么', '什么是', '什么意思', '指的是', '啥意思', '意思'];
    for (var a = 0; a < askWords.length; a++) {
      if (lower.indexOf(askWords[a].toLowerCase()) !== -1) {
        for (var k = 0; k < conceptKeys.length; k++) {
          if (lower.indexOf(conceptKeys[k].toLowerCase()) !== -1) return { type: 'concept' };
        }
      }
    }

    // 3) 无明确疑问词但含环保概念词（如「塑料污染」「碳足迹」），纯商品名不触发
    var isPureProduct = matchAny(MY_ECO_ALIASES, input) || matchAny(MY_NORMAL_ALIASES, input);
    if (!isPureProduct) {
      for (var m = 0; m < conceptKeys.length; m++) {
        if (lower.indexOf(conceptKeys[m].toLowerCase()) !== -1) return { type: 'concept' };
      }
    }

    // 4) 含疑问语气但未命中任何词库 -> 标记为未知追问（走兜底引导）
    var askMarkers = ['?', '？', '吗', '呢', '什么', '怎么', '如何', '多少', '哪个'];
    for (var n = 0; n < askMarkers.length; n++) {
      if (lower.indexOf(askMarkers[n].toLowerCase()) !== -1) return { type: 'unknown' };
    }
    return null;
  }

  function detectTarget(input) {
    if (matchAny(B_TARGET_MARKERS, input)) return 'B';
    if (matchAny(A_TARGET_MARKERS, input)) return 'A';
    return 'both';
  }

  function answerQuestion(qt, input) {
    if (qt.type === '帮助') {
      return HELP_MSG;
    }
    if (qt.type === '对比') {
      return KB_COMPARE + getClosing('对比');
    }
    if (qt.type === '推荐') {
      return KB_RECOMMEND + getClosing('推荐');
    }
    if (qt.type === 'concept') {
      var lower2 = input.trim().toLowerCase();
      for (var key in ECO_CONCEPT_KEYWORDS) {
        if (lower2.indexOf(key.toLowerCase()) !== -1) {
          return ECO_CONCEPTS[ECO_CONCEPT_KEYWORDS[key]] + getClosing('eco');
        }
      }
      return ECO_CONCEPTS['绿色消费'] + getClosing('eco');
    }

    var label = LABEL_MAP[qt.type];
    var target = detectTarget(input);
    if (target === 'A') {
      return 'A（' + MY_NORMAL_PRODUCT + '）的' + label + '：' + KB_A[qt.type] + getClosing(qt.type);
    }
    if (target === 'B') {
      return 'B（' + MY_ECO_PRODUCT + '）的' + label + '：' + KB_B[qt.type] + getClosing(qt.type);
    }
    return 'A（' + MY_NORMAL_PRODUCT + '）的' + label + '：' + KB_A[qt.type]
      + 'B（' + MY_ECO_PRODUCT + '）的' + label + '：' + KB_B[qt.type]
      + getClosing(qt.type);
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
    var nextStage = flowStage;

    // 阶段0：等待被试输入普通商品 -> 展示核心操纵文案（干预文案）
    if (flowStage === 0) {
      if (matchAny(MY_NORMAL_ALIASES, text)) {
        reply = MY_COPY;
        replyType = 'product_copy';
        nextStage = 1;
        setPlaceholder(1);
      } else if (matchAny(MY_ECO_ALIASES, text)) {
        reply = ECO_PROMPT_MSG;
        replyType = 'eco_product_prompt';
      } else {
        reply = ERROR_MSG;
        replyType = 'error_response';
      }
    }
    // 阶段1：已展示干预文案，引导询问本店商品 -> 展示 A/B 外观+价格
    else if (flowStage === 1) {
      if (isAskProduct(text)) {
        reply = PRODUCT_INFO;
        replyType = 'product_info';
        nextStage = 2;
        setPlaceholder(2);
      } else {
        // 尚未询问商品，继续引导
        reply = '您还没有询问本店可购买的商品哦，请输入"本店有什么商品"或商品名称，我来为您介绍。';
        replyType = 'guidance';
      }
    }
    // 阶段2：已展示商品信息，进入自由问答
    else {
      var qt = detectQuestion(text);
      if (qt) {
        if (qt.type === 'unknown') {
          reply = FALLBACK_QA_MSG;
          replyType = 'fallback_qa';
        } else {
          reply = answerQuestion(qt, text);
          replyType = 'qa_response';
        }
      }
      // 被试再次输入环保款商品 -> 提示改输普通版本
      else if (matchAny(MY_ECO_ALIASES, text)) {
        reply = ECO_PROMPT_MSG;
        replyType = 'eco_product_prompt';
      }
      // 被试再次询问商品信息 -> 重新展示 A/B 外观+价格
      else if (isAskProduct(text)) {
        reply = PRODUCT_INFO;
        replyType = 'product_info';
      }
      // 再次输入普通商品 -> 重新展示干预文案
      else if (matchAny(MY_NORMAL_ALIASES, text)) {
        reply = MY_COPY;
        replyType = 'product_copy';
      }
      // 其余输入 -> 按品类固定报错
      else {
        reply = ERROR_MSG;
        replyType = 'error_response';
      }

      // 词库未覆盖的追问（含疑问语气但未命中任何词库）-> 引导可询问内容
      if (replyType === 'error_response') {
        var askMarkers = ['?', '？', '吗', '呢', '么', '什么', '怎么', '如何', '多少', '哪个', '哪'];
        for (var ai = 0; ai < askMarkers.length; ai++) {
          if (text.indexOf(askMarkers[ai]) !== -1) {
            reply = FALLBACK_QA_MSG;
            replyType = 'fallback_qa';
            break;
          }
        }
      }
    }

    flowStage = nextStage;

    sendBtn.disabled = true;
    setTimeout(function() {
      replaceWithSystemMessage(loadingDiv, reply);
      reportToCredamo(replyType, { input: text, stage: flowStage });
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


def build_product_info(cat):
    """阶段2：展示 A/B 两款商品的外观与价格信息。"""
    return (
        '本店有两种%s售卖：\n\n'
        'A（非环保产品）%s\n'
        '%s\n'
        '价格：%s元\n\n'
        'B（环保产品）%s\n'
        '%s\n'
        '价格：%s元' % (
            cat['measure'],
            cat['normal_name'], cat['normal_appearance'], cat['normal_price'],
            cat['eco_name'], cat['eco_appearance'], cat['eco_price'],
        )
    )


def generate_html(group_type, cat):
    group_label = 'collaboration' if group_type == 'collaboration' else 'non_collaboration'
    copy_dict = COLLABORATION_COPY if group_type == 'collaboration' else NON_COLLABORATION_COPY
    copy = copy_dict[cat['copy_key']]
    product_info = build_product_info(cat)

    # 合并基础知识库（7 项）与扩展知识库（10 项）
    extra = EXTRA_KB[cat['category']]
    kb_a = dict(cat['normal_kb'])
    kb_a.update(extra['normal'])
    kb_b = dict(cat['eco_kb'])
    kb_b.update(extra['eco'])

    js = (JS_TEMPLATE
          .replace('__CATEGORY__', json.dumps(cat['category'], ensure_ascii=False))
          .replace('__MEASURE__', json.dumps(cat['measure'], ensure_ascii=False))
          .replace('__NORMAL_NAME__', json.dumps(cat['normal_name'], ensure_ascii=False))
          .replace('__ECO_NAME__', json.dumps(cat['eco_name'], ensure_ascii=False))
          .replace('__GROUP_LABEL__', json.dumps(group_label, ensure_ascii=False))
          .replace('__NORMAL_ALIASES__', json.dumps(cat['normal_aliases'], ensure_ascii=False))
          .replace('__ECO_ALIASES__', json.dumps(cat['eco_aliases'], ensure_ascii=False))
          .replace('__COPY__', json.dumps(copy, ensure_ascii=False))
          .replace('__PRODUCT_INFO__', json.dumps(product_info, ensure_ascii=False))
          .replace('__ECO_CONCEPTS__', json.dumps(ECO_CONCEPTS, ensure_ascii=False))
          .replace('__ECO_CONCEPT_KEYWORDS__', json.dumps(ECO_CONCEPT_KEYWORDS, ensure_ascii=False))
          .replace('__KB_A__', json.dumps(kb_a, ensure_ascii=False))
          .replace('__KB_B__', json.dumps(kb_b, ensure_ascii=False))
          .replace('__KB_COMPARE__', json.dumps(cat['compare'], ensure_ascii=False))
          .replace('__KB_RECOMMEND__', json.dumps(cat['recommend'], ensure_ascii=False)))

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
            '    <input type="text" id="userInput" placeholder="请输入产品名称" autocomplete="off" enterkeyhint="send">\n'
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
            filename = '%02d_%s_%s.html' % (count, group_label, cat['category'])
            filepath = os.path.join(OUT_DIR, filename)

            html = generate_html(group_type, cat)
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(html)

            print('[%02d/12] %s  (%d bytes)' % (count, filename, len(html.encode('utf-8'))))

    print('\n全部 12 个 Agent HTML 文件已生成到: %s' % OUT_DIR)


if __name__ == '__main__':
    main()
