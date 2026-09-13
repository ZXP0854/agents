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
  var NORMAL_RESPONSE = __NORMAL_RESPONSE__;

  /* ========== A/B 知识库 ========== */
  var KB_A = __KB_A__;
  var KB_B = __KB_B__;
  var KB_COMPARE = __KB_COMPARE__;
  var KB_RECOMMEND = __KB_RECOMMEND__;

  /* ========== 固定话术（实验要求，不可改动） ========== */
  var WELCOME_MSG = '顾客您好！欢迎光临小林超市，我是本店的 AI 智能助手小林，请输入您的商品信息。';
  var ERROR_MSG = '抱歉，暂不支持该商品，请输入与' + MY_MEASURE + '有关的商品。';
  var ECO_PROMPT_MSG = '本次请输入普通版本商品，请重新输入对应普通商品名称。';

  /* ========== 组别结尾语（协作/非协作关键词区分） ========== */
  var GROUP_CLOSING = (MY_GROUP === 'collaboration')
    ? '加入我们，让我们一起选择' + MY_ECO_PRODUCT + '，为环保做出贡献！'
    : '请您选择' + MY_ECO_PRODUCT + '，为环保做出贡献！';

  /* ========== 追问类型（12 个 agent 风格统一） ========== */
  var QUESTION_TYPES = [
    { type: '推荐', target: false, keys: ['推荐', '建议', '买哪个', '选哪个', '应该买', '值得买', '买什么', '怎么选', '买哪款', '选哪款'] },
    { type: '对比', target: false, keys: ['对比', '区别', '比较', '相比', 'vs', '哪个好', '有什么不同', '差别', '不一样', '哪个更好', '不同'] },
    { type: '优点', target: true,  keys: ['优点', '优势', '特色', '亮点', '哪里好', '好在', '好不好', '怎么样'] },
    { type: '好处', target: true,  keys: ['好处', '有什么用', '有什么好', '作用', '帮助'] },
    { type: '缺点', target: true,  keys: ['缺点', '不足', '劣势', '缺陷', '坏处', '不好', '问题'] },
    { type: '用途', target: true,  keys: ['用途', '怎么用', '使用场景', '能做什么', '适用', '适合', '场景', '干什么用', '干嘛用'] },
    { type: '价格', target: true,  keys: ['价格', '多少钱', '贵', '便宜', '成本', '花费', '划算', '价位'] },
    { type: '安全', target: true,  keys: ['安全', '有毒', '无害', '健康', '放心', '标准', '检测', '危害'] },
    { type: '耐用', target: true,  keys: ['耐用', '质量', '寿命', '持久', '容易坏', '结实', '保存', '保质', '储存', '好用吗', '能用多久'] }
  ];

  var LABEL_MAP = { '优点': '优点', '好处': '好处', '缺点': '缺点', '用途': '用途', '价格': '价格', '安全': '安全性', '耐用': '耐用性' };

  var A_TARGET_MARKERS = MY_NORMAL_ALIASES.concat(['普通款', '普通', 'a款', '选a', '买a', 'a的', 'a有', 'a好', 'a是', 'a什么']);
  var B_TARGET_MARKERS = MY_ECO_ALIASES.concat(['环保款', '可降解款', '绿色款', '环保', '可降解', 'b款', '选b', '买b', 'b的', 'b有', 'b好', 'b是', 'b什么']);

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

  /* ========== 追问检测 ========== */
  function detectQuestion(input) {
    var lower = input.trim().toLowerCase();
    for (var i = 0; i < QUESTION_TYPES.length; i++) {
      var qt = QUESTION_TYPES[i];
      for (var j = 0; j < qt.keys.length; j++) {
        if (lower.indexOf(qt.keys[j].toLowerCase()) !== -1) return qt;
      }
    }
    // 环保概念追问兜底（如「可降解是什么」「B环保吗」），但输入本身是商品名时不触发
    if (matchAny(MY_ECO_ALIASES, input) || matchAny(MY_NORMAL_ALIASES, input)) return null;
    var ecoConceptKeys = ['环保', '可降解', '降解', '绿色', '可持续'];
    for (var k = 0; k < ecoConceptKeys.length; k++) {
      if (lower.indexOf(ecoConceptKeys[k].toLowerCase()) !== -1) return { type: 'eco' };
    }
    return null;
  }

  function detectTarget(input) {
    if (matchAny(B_TARGET_MARKERS, input)) return 'B';
    if (matchAny(A_TARGET_MARKERS, input)) return 'A';
    return 'both';
  }

  function answerQuestion(qt, input) {
    if (qt.type === '对比') {
      return KB_COMPARE + '\\n' + GROUP_CLOSING;
    }
    if (qt.type === '推荐') {
      return KB_RECOMMEND + '\\n' + GROUP_CLOSING;
    }
    if (qt.type === 'eco') {
      return 'B（' + MY_ECO_PRODUCT + '）的好处：' + KB_B['好处'] + '\\n' + GROUP_CLOSING;
    }

    var label = LABEL_MAP[qt.type];
    var target = detectTarget(input);
    if (target === 'A') {
      return 'A（' + MY_NORMAL_PRODUCT + '）的' + label + '：' + KB_A[qt.type] + '\\n' + GROUP_CLOSING;
    }
    if (target === 'B') {
      return 'B（' + MY_ECO_PRODUCT + '）的' + label + '：' + KB_B[qt.type] + '\\n' + GROUP_CLOSING;
    }
    return 'A（' + MY_NORMAL_PRODUCT + '）的' + label + '：' + KB_A[qt.type]
      + '\\nB（' + MY_ECO_PRODUCT + '）的' + label + '：' + KB_B[qt.type]
      + '\\n' + GROUP_CLOSING;
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

    // 1) 追问（优点/好处/缺点/用途/价格/安全/耐用/对比/推荐/环保概念）
    var qt = detectQuestion(text);
    if (qt) {
      reply = answerQuestion(qt, text);
      replyType = 'qa_response';
    }
    // 2) 被试直接输入环保款商品 -> 提示改输普通版本
    else if (matchAny(MY_ECO_ALIASES, text)) {
      reply = ECO_PROMPT_MSG;
      replyType = 'eco_product_prompt';
    }
    // 3) 输入本品类普通商品 -> 商品信息 + 干预文案
    else if (matchAny(MY_NORMAL_ALIASES, text)) {
      reply = NORMAL_RESPONSE;
      replyType = 'product_response';
    }
    // 4) 其余输入 -> 按品类固定报错
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
    """普通款输入 -> 「本店有两种X售卖：」+ 普通款/环保款外观价格 + 干预文案。"""
    return (
        '本店有两种%s售卖：\n\n'
        'A（普通款）%s\n'
        '%s\n'
        '价格：%s元\n\n'
        'B（环保可降解款）%s\n'
        '%s\n'
        '价格：%s元\n\n'
        '%s' % (
            cat['measure'],
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
          .replace('__MEASURE__', json.dumps(cat['measure'], ensure_ascii=False))
          .replace('__NORMAL_NAME__', json.dumps(cat['normal_name'], ensure_ascii=False))
          .replace('__ECO_NAME__', json.dumps(cat['eco_name'], ensure_ascii=False))
          .replace('__GROUP_LABEL__', json.dumps(group_label, ensure_ascii=False))
          .replace('__NORMAL_ALIASES__', json.dumps(cat['normal_aliases'], ensure_ascii=False))
          .replace('__ECO_ALIASES__', json.dumps(cat['eco_aliases'], ensure_ascii=False))
          .replace('__NORMAL_RESPONSE__', json.dumps(normal_response, ensure_ascii=False))
          .replace('__KB_A__', json.dumps(cat['normal_kb'], ensure_ascii=False))
          .replace('__KB_B__', json.dumps(cat['eco_kb'], ensure_ascii=False))
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
            filename = '%02d_%s_%s.html' % (count, group_label, cat['category'])
            filepath = os.path.join(OUT_DIR, filename)

            html = generate_html(group_type, cat)
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(html)

            print('[%02d/12] %s  (%d bytes)' % (count, filename, len(html.encode('utf-8'))))

    print('\n全部 12 个 Agent HTML 文件已生成到: %s' % OUT_DIR)


if __name__ == '__main__':
    main()
