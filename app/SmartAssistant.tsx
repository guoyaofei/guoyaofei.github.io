"use client";

import { FormEvent, useEffect, useRef, useState } from "react";

type Reply = {
  title: string;
  body: string;
  bullets?: string[];
  source: string;
  action?: { label: string; href: string };
};

type Message = {
  id: number;
  from: "assistant" | "visitor";
  text?: string;
  reply?: Reply;
};

type Diagnosis = {
  role?: string;
  need?: string;
  stage?: string;
};

const welcomeReply: Reply = {
  title: "您好，我是早舟智能咨询助手",
  body: "我可以帮助您了解早舟科技、产业学院、学校内涵建设、AI赋能场景和平哥系列产品，也可以用三个问题完成一次需求预诊断。",
  bullets: ["了解三大业务与产品体系", "快速匹配适合的服务方向", "完成三步需求预诊断"],
  source: "早舟官网公开信息",
};

const quickQuestions = [
  "我来自院校",
  "我来自企业",
  "我想了解课程",
  "我有项目难题",
];

const humanReply: Reply = {
  title: "这类问题建议由工作人员确认",
  body: "课程报名、工作坊与日常服务预约可联系朱老师；机构项目、专家咨询、战略合作及复杂项目可联系杜老师。发送完整咨询摘要，可以减少重复沟通。",
  bullets: ["朱老师：152 1321 7592", "杜老师：186 9698 7413", "合同、退款、发票、投诉与最终报价必须由人工确认"],
  source: "官网联系我们与服务规则",
  action: { label: "查看联系方式与二维码", href: "#contact" },
};

function knowledgeReply(question: string): Reply {
  const q = question.toLowerCase().replaceAll("，", "").replaceAll("？", "").trim();
  const has = (...words: string[]) => words.some((word) => q.includes(word));

  if (has("退款", "合同", "发票", "投诉", "优惠", "折扣", "保证", "承诺", "一定成功", "包过", "获奖", "验收")) {
    return humanReply;
  }

  if (has("人工", "联系", "电话", "微信", "报名", "预约")) {
    return humanReply;
  }

  if (has("399", "成长圈")) {
    return {
      title: "平哥产教融合成长圈",
      body: "公开参考价为399元/年，面向院校干部、教师、企业项目人员和职业教育同行。核心价值是持续获得行业判断、案例会诊、工具资料与社群连接。",
      bullets: ["全年约40场线上成长夜话", "政策解读、案例拆解与公开会诊", "不包含一对一咨询、方案代写和无限答疑"],
      source: "官网第04部分：产教融合产品矩阵",
      action: { label: "查看成长圈详情", href: "#product-2" },
    };
  }

  if (has("8800", "专题班", "专题工作坊")) {
    return {
      title: "产教融合专题实战工作坊",
      body: "公开参考价为8800元，采用两天一夜、小班共创方式，围绕一个明确问题集中突破，适合已有单点需求的客户。",
      bullets: ["AI赋能产教融合", "企业进入职业院校", "产业学院运营与项目设计", "职教专家IP与知识产品化"],
      source: "官网第04部分：产教融合产品矩阵",
      action: { label: "查看专题工作坊", href: "#product-3" },
    };
  }

  if (has("19800", "旗舰", "项目共创")) {
    return {
      title: "项目操盘与成果共创旗舰工作坊",
      body: "公开参考价为19800元，适合正在承担真实项目的负责人。重点不是多听课程，而是形成诊断报告、合作机制、成果路径与90天行动清单。",
      bullets: ["课前项目诊断", "三天两夜现场共创", "一对一专项咨询", "阶段性跟踪"],
      source: "官网第04部分：产教融合产品矩阵",
      action: { label: "查看旗舰工作坊", href: "#product-4" },
    };
  }

  if (has("39800", "操盘手", "研修计划")) {
    return {
      title: "产教融合项目操盘手研修计划",
      body: "公开参考价为39800元，采用申请制和3至6个月周期式培养，目标是从听懂产教融合升级为能够独立操盘项目。",
      bullets: ["战略判断与项目设计", "方案写作与商务谈判", "真实项目与导师辅导", "项目路演和个人案例"],
      source: "官网第04部分：产教融合产品矩阵",
      action: { label: "查看操盘手计划", href: "#product-5" },
    };
  }

  if (has("10万", "100000", "大会员", "事业协同")) {
    return {
      title: "平哥产教融合事业大会员",
      body: "公开参考价为10万元/年，采用邀请审核制。购买的不是更多课程，而是年度战略顾问、重大项目会诊与事业协同服务。",
      bullets: ["年度战略诊断", "固定次数一对一咨询", "重大项目会诊", "闭门共创与年度复盘"],
      source: "官网第04部分：产教融合产品矩阵",
      action: { label: "查看事业大会员", href: "#product-6" },
    };
  }

  if (has("机构", "定制", "学校项目", "企业项目")) {
    return {
      title: "机构定制解决方案",
      body: "面向院校、二级学院、央国企、产业园区和行业组织，参考价格10万元起，实际报价需根据项目范围、周期、调研深度和交付成果由人工确认。",
      bullets: ["团队调研与项目诊断", "顶层设计和产业学院运营", "90天或年度项目陪跑", "成果提炼与标杆案例打造"],
      source: "官网第04部分：机构定制解决方案",
      action: { label: "填写机构项目需求", href: "#diagnosis" },
    };
  }

  if (has("多少钱", "价格", "费用", "收费", "产品")) {
    return {
      title: "早舟公开参考产品路径",
      body: "目前官网展示免费内容、399元成长圈、8800元专题工作坊、19800元旗舰工作坊、39800元操盘手计划、10万元事业大会员及10万元起的机构定制服务。",
      bullets: ["公开价格用于帮助理解产品层级", "名额、权益、发票与履约规则以最终说明为准", "机构项目和复杂需求需要单独诊断报价"],
      source: "官网第04部分：产教融合产品矩阵",
      action: { label: "查看完整产品矩阵", href: "#products" },
    };
  }

  if (has("产业学院", "产教融合")) {
    return {
      title: "产教融合与产业学院服务",
      body: "早舟从合作模式、专业群与产业链对接开始，进一步支持产业学院实体化运营、企业参与机制、知识资产化和成果评价，目标是让项目持续运行而不止于签约挂牌。",
      bullets: ["产业学院规划与运营", "产教融合项目策划", "校企合作机制建设", "成果提炼与标杆案例打造"],
      source: "官网第02部分：三大业务",
      action: { label: "查看解决方案", href: "#solutions" },
    };
  }

  if (has("学校", "内涵", "党建", "德育", "教材", "课题")) {
    return {
      title: "学校内涵建设服务",
      body: "围绕党建、战略规划、教学、德育和校园文化，早舟提供从调研、规划、计划和课题研究，到成果打造、经验总结、教师培训与教材编制的全过程支持。",
      bullets: ["战略规划与专项计划", "课题研究与成果表达", "文化、教学与德育策划", "教师培训与教材编制"],
      source: "官网第02部分：三大业务",
      action: { label: "填写学校需求", href: "#diagnosis" },
    };
  }

  if (has("ai", "人工智能", "工作流", "智能")) {
    return {
      title: "AI赋能教育与产教融合场景",
      body: "早舟主张让AI进入真实工作流，而不是停留在工具演示。典型场景包括调研分析、方案策划、教学应用、项目运营、成果沉淀与团队培训。",
      bullets: ["需求诊断", "智能策划", "实体运营", "成果量化"],
      source: "官网第03部分：AI工作流",
      action: { label: "查看AI工作流", href: "#workflow" },
    };
  }

  if (has("平哥", "杜平", "创始人", "董事长")) {
    return {
      title: "杜平（平哥）",
      body: "杜平是早舟科技创始人兼董事长，长期深耕职业教育与产教融合，带领团队深入学校和企业现场，关注产业学院运营、学校内生能力和AI工作流。",
      bullets: ["职业教育与产教融合实战专家", "平哥聊职教内容品牌创始人", "多家国企、央企及民企产教融合顾问", "重庆市西部产教融合研究院研究员、首席顾问"],
      source: "官网第01部分：关于早舟",
      action: { label: "查看平哥介绍", href: "#about" },
    };
  }

  if (has("案例", "成果", "客户", "合作过")) {
    return {
      title: "案例成果与公开依据",
      body: "官网已展示产教融合线下工作坊、产业学院实体化运营辅导、AI时代学校管理能力提升等代表案例，并整理了更多院校和政府网站公开报道。",
      bullets: ["每个案例优先链接原发布单位", "更多二十年案例将根据客户授权持续上线", "复杂项目可由工作人员匹配同类经验"],
      source: "官网第05部分：案例成果",
      action: { label: "查看案例成果", href: "#cases" },
    };
  }

  if (has("视频", "直播", "公众号", "文章", "内容")) {
    return {
      title: "平哥内容中心",
      body: "官网内容中心已汇集平哥的两个视频号、两个公众号及专题资料库，可直接扫描对应二维码关注，持续获取视频、直播、文章与行业洞察。",
      bullets: ["政策与趋势", "产业学院", "AI赋能", "项目操盘"],
      source: "官网第06部分：内容中心",
      action: { label: "进入内容中心", href: "#media" },
    };
  }

  if (has("公司", "早舟", "九源", "做什么", "介绍")) {
    return {
      title: "重庆市早舟科技有限公司",
      body: "早舟科技前身为九源教育，团队深耕职业教育20年，主要服务全国大中专院校、本科院校，以及产教融合相关国企、央企和优质民营企业。",
      bullets: ["产教融合与产业学院", "学校内涵建设", "AI赋能教育与产教融合工作场景"],
      source: "官网第01—02部分：关于早舟与三大业务",
      action: { label: "了解早舟科技", href: "#about" },
    };
  }

  if (has("院校", "老师", "教师", "学校领导")) {
    return {
      title: "院校客户可以从这里开始",
      body: "建议先明确您所在学校、负责岗位、项目阶段和最希望解决的问题。若只是持续学习，可先了解成长圈；已有具体任务，可进入专题工作坊或项目诊断。",
      bullets: ["初步了解：内容中心或399元成长圈", "单点难题：8800元专题工作坊", "真实复杂项目：19800元旗舰班或机构服务"],
      source: "官网产品匹配规则",
      action: { label: "开始完整需求预诊断", href: "#diagnosis" },
    };
  }

  if (has("企业", "央企", "国企", "民企", "入校")) {
    return {
      title: "企业客户可以从这里开始",
      body: "企业进入职业院校通常需要先厘清资源、产品、目标专业、合作机制和持续收益。单点能力提升可选择专题工作坊，真实入校项目建议先做机构诊断。",
      bullets: ["企业入校路径设计", "校企合作机制与商务模式", "产业学院或专业群项目", "知识产品化与标杆成果"],
      source: "官网企业客户匹配规则",
      action: { label: "提交企业项目需求", href: "#diagnosis" },
    };
  }

  return {
    title: "这个问题需要补充一些信息",
    body: "当前官网公开资料还不足以可靠回答这个问题。您可以换一种说法，或填写需求预诊断，由工作人员结合单位、身份和项目阶段进一步判断。",
    bullets: ["不会编造没有依据的答案", "复杂项目、合同和最终价格由人工确认", "可填写需求预诊断获得进一步判断"],
    source: "官网安全回答规则",
    action: { label: "填写需求预诊断", href: "#diagnosis" },
  };
}

function diagnosisReply(profile: Diagnosis): Reply {
  const { role = "暂未选择", need = "综合需求", stage = "初步了解" } = profile;
  let recommendation = "建议先查看免费内容与399元成长圈，持续了解趋势和方法，再决定是否进入线下产品。";
  let href = "#product-2";
  let label = "查看成长圈";

  if (need === "课程与能力成长") {
    recommendation = stage === "初步了解"
      ? "建议先从399元成长圈开始；如果已经有明确单点任务，可进一步了解8800元专题工作坊。"
      : "建议重点了解8800元专题工作坊或39800元操盘手研修计划，并由工作人员结合您的真实任务进行匹配。";
    href = "#products";
    label = "查看产品矩阵";
  } else if (need === "产业学院与产教融合") {
    recommendation = stage === "推进受阻" || stage === "需要成果提炼"
      ? "您更接近真实复杂项目，建议优先做完整项目诊断，再判断19800元旗舰工作坊或机构定制服务。"
      : "建议先明确合作企业、目标专业、当前机制和预期成果，再判断8800元专题工作坊或项目共创产品。";
    href = "#diagnosis";
    label = "填写完整项目信息";
  } else if (need === "学校内涵建设") {
    recommendation = "学校内涵建设通常涉及多部门与成果体系，建议直接填写完整需求，由工作人员判断是专项策划、教师培训还是机构项目服务。";
    href = "#diagnosis";
    label = "填写学校需求";
  } else if (need === "AI赋能工作流") {
    recommendation = stage === "初步了解"
      ? "建议先查看AI工作流和相关内容；如已有明确场景，可考虑AI赋能产教融合专题工作坊。"
      : "建议把当前流程、使用人员、现有工具和希望形成的成果写入完整诊断，由早舟匹配培训或项目共创方案。";
    href = "#workflow";
    label = "查看AI工作流";
  }

  return {
    title: "需求预诊断已生成",
    body: recommendation,
    bullets: [`您的身份：${role}`, `关注方向：${need}`, `项目阶段：${stage}`, "最终方案、报价与项目安排以人工确认为准"],
    source: "官网需求匹配规则",
    action: { label, href },
  };
}

export default function SmartAssistant() {
  const [open, setOpen] = useState(false);
  const [messages, setMessages] = useState<Message[]>([{ id: 1, from: "assistant", reply: welcomeReply }]);
  const [input, setInput] = useState("");
  const [diagnosis, setDiagnosis] = useState<Diagnosis>({});
  const [diagnosisStep, setDiagnosisStep] = useState(0);
  const messageEndRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (open) messageEndRef.current?.scrollIntoView({ behavior: "smooth", block: "nearest" });
  }, [messages, open, diagnosisStep]);

  function addExchange(question: string, reply = knowledgeReply(question)) {
    const stamp = Date.now();
    setMessages((current) => [
      ...current,
      { id: stamp, from: "visitor", text: question },
      { id: stamp + 1, from: "assistant", reply },
    ]);
  }

  function submitQuestion(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const question = input.trim();
    if (!question) return;
    addExchange(question);
    setInput("");
  }

  function chooseQuick(question: string) {
    if (question === "我有项目难题") {
      addExchange(question, {
        title: "先用三个问题做一次需求预诊断",
        body: "我会根据您的身份、关注方向和项目阶段给出初步路径。",
        source: "官网需求匹配规则",
      });
      setDiagnosis({});
      setDiagnosisStep(1);
      return;
    }
    addExchange(question);
  }

  function chooseDiagnosis(value: string) {
    if (diagnosisStep === 1) {
      setDiagnosis({ role: value });
      setDiagnosisStep(2);
      return;
    }
    if (diagnosisStep === 2) {
      setDiagnosis((current) => ({ ...current, need: value }));
      setDiagnosisStep(3);
      return;
    }
    const completed = { ...diagnosis, stage: value };
    setDiagnosis(completed);
    setDiagnosisStep(0);
    addExchange("完成三步需求预诊断", diagnosisReply(completed));
  }

  function resetConversation() {
    setMessages([{ id: Date.now(), from: "assistant", reply: welcomeReply }]);
    setDiagnosis({});
    setDiagnosisStep(0);
    setInput("");
  }

  const diagnosisOptions = diagnosisStep === 1
    ? ["院校管理者", "教师/项目负责人", "企业负责人", "其他同行"]
    : diagnosisStep === 2
      ? ["产业学院与产教融合", "学校内涵建设", "AI赋能工作流", "课程与能力成长"]
      : ["初步了解", "正在立项", "推进受阻", "需要成果提炼"];

  return (
    <aside className={`smart-assistant${open ? " is-open" : ""}`} id="ai-consulting" aria-label="早舟智能咨询助手">
      <button className="assistant-launcher" type="button" onClick={() => setOpen(true)} aria-haspopup="dialog" aria-expanded={open}>
        <span className="assistant-launcher-icon">AI</span>
        <span><small>在线咨询</small><strong>智能咨询</strong></span>
      </button>

      {open && <div className="assistant-backdrop" onClick={() => setOpen(false)} aria-hidden="true" />}

      <section className="assistant-panel" role="dialog" aria-modal="false" aria-labelledby="assistant-title" hidden={!open}>
        <header className="assistant-panel-header">
          <div>
            <span className="assistant-status"><i /> 在线咨询</span>
            <h2 id="assistant-title">早舟智能咨询助手</h2>
          </div>
          <button type="button" onClick={() => setOpen(false)} aria-label="关闭智能咨询助手">×</button>
        </header>

        <div className="assistant-notice">
          <strong>官网智能导览</strong>
          <span>重要价格、合同与项目安排由工作人员最终确认</span>
        </div>

        <div className="assistant-messages" aria-live="polite">
          {messages.map((message) => message.from === "visitor" ? (
            <div className="assistant-message visitor-message" key={message.id}>{message.text}</div>
          ) : (
            <article className="assistant-message assistant-reply" key={message.id}>
              <span className="assistant-avatar">早舟AI</span>
              <div>
                <h3>{message.reply?.title}</h3>
                <p>{message.reply?.body}</p>
                {message.reply?.bullets && <ul>{message.reply.bullets.map((bullet) => <li key={bullet}>{bullet}</li>)}</ul>}
                <small>回答依据：{message.reply?.source}</small>
                {message.reply?.action && <a href={message.reply.action.href} onClick={() => setOpen(false)}>{message.reply.action.label} <span>→</span></a>}
              </div>
            </article>
          ))}

          {diagnosisStep > 0 && (
            <div className="assistant-diagnosis-step">
              <span>需求预诊断 · 第 {diagnosisStep} / 3 步</span>
              <h3>{diagnosisStep === 1 ? "您的身份更接近哪一类？" : diagnosisStep === 2 ? "这次最关注哪个方向？" : "项目目前处于什么阶段？"}</h3>
              <div>{diagnosisOptions.map((option) => <button type="button" key={option} onClick={() => chooseDiagnosis(option)}>{option}</button>)}</div>
              <button className="assistant-cancel-diagnosis" type="button" onClick={() => setDiagnosisStep(0)}>暂不诊断</button>
            </div>
          )}
          <div ref={messageEndRef} />
        </div>

        {diagnosisStep === 0 && (
          <div className="assistant-quick-questions" aria-label="常用问题">
            {quickQuestions.map((question) => <button type="button" key={question} onClick={() => chooseQuick(question)}>{question}</button>)}
          </div>
        )}

        <form className="assistant-input" onSubmit={submitQuestion}>
          <label className="sr-only" htmlFor="assistant-question">输入您的问题</label>
          <input id="assistant-question" value={input} onChange={(event) => setInput(event.target.value)} placeholder="例如：8800元工作坊适合谁？" maxLength={160} />
          <button type="submit" aria-label="发送问题">↑</button>
        </form>

        <footer className="assistant-panel-footer">
          <button type="button" onClick={resetConversation}>重新开始</button>
          <a href="#contact" onClick={() => setOpen(false)}>转人工服务</a>
        </footer>
      </section>
    </aside>
  );
}
