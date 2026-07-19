"use client";

import { FormEvent, useState } from "react";

const field = (form: FormData, name: string) => String(form.get(name) || "未填写").trim() || "未填写";

export default function ConsultationForm() {
  const [summary, setSummary] = useState("");
  const [copied, setCopied] = useState(false);

  function generateSummary(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const data = new FormData(event.currentTarget);
    setSummary([
      "【早舟项目咨询预诊断】",
      `姓名：${field(data, "name")}`,
      `单位：${field(data, "organization")}`,
      `身份：${field(data, "role")}`,
      `联系方式：${field(data, "contact")}`,
      `需求方向：${field(data, "need")}`,
      `项目阶段：${field(data, "stage")}`,
      `计划时间：${field(data, "timeline")}`,
      `预算范围：${field(data, "budget")}`,
      `核心问题：${field(data, "problem")}`,
    ].join("\n"));
    setCopied(false);
  }

  async function copySummary() {
    if (!summary) return;
    await navigator.clipboard.writeText(summary);
    setCopied(true);
  }

  return (
    <section className="diagnosis-section" id="diagnosis">
      <div className="section diagnosis-inner">
        <div className="diagnosis-intro">
          <p className="section-index light">07 / 需求预诊断</p>
          <h2>先把基础情况说清楚，<br />再进入有效沟通</h2>
          <p>用约 2 分钟填写项目基础信息，系统会整理成一份咨询摘要。复制后发送给朱老师或杜老师，工作人员可直接判断适合的产品、专家和下一步安排。</p>
          <ol>
            <li><span>01</span>填写单位、身份与需求</li>
            <li><span>02</span>生成结构化咨询摘要</li>
            <li><span>03</span>复制并扫码发送给工作人员</li>
          </ol>
        </div>
        <form className="diagnosis-form" onSubmit={generateSummary}>
          <div className="form-grid">
            <label>姓名<span>*</span><input name="name" required autoComplete="name" /></label>
            <label>单位/学校<span>*</span><input name="organization" required autoComplete="organization" /></label>
            <label>您的身份<span>*</span><select name="role" required defaultValue=""><option value="" disabled>请选择</option><option>学校领导</option><option>二级学院负责人</option><option>教师/项目负责人</option><option>企业负责人</option><option>企业项目人员</option><option>其他</option></select></label>
            <label>手机或微信<span>*</span><input name="contact" required inputMode="tel" autoComplete="tel" /></label>
            <label>需求方向<span>*</span><select name="need" required defaultValue=""><option value="" disabled>请选择</option><option>产业学院建设与运营</option><option>产教融合项目策划</option><option>学校内涵建设</option><option>AI教育/产教融合工作流</option><option>工作坊、课程或成长圈</option><option>长期顾问与事业协同</option><option>其他</option></select></label>
            <label>目前阶段<select name="stage" defaultValue="初步了解"><option>初步了解</option><option>正在立项</option><option>已经推进，遇到困难</option><option>需要成果提炼</option><option>寻找长期合作</option></select></label>
            <label>计划启动时间<select name="timeline" defaultValue="1—3个月"><option>1个月内</option><option>1—3个月</option><option>3—6个月</option><option>6个月以后</option><option>暂未确定</option></select></label>
            <label>预算范围（选填）<select name="budget" defaultValue="暂未确定"><option>暂未确定</option><option>1万元以内</option><option>1—5万元</option><option>5—10万元</option><option>10—30万元</option><option>30万元以上</option></select></label>
            <label className="form-span">最希望解决的一个问题<textarea name="problem" rows={4} placeholder="例如：产业学院已挂牌，但企业参与不足、项目成果难以量化……" /></label>
          </div>
          <button className="button button-primary" type="submit">生成咨询摘要 <span>→</span></button>
          {summary && <div className="form-summary" aria-live="polite"><pre>{summary}</pre><button type="button" onClick={copySummary}>{copied ? "已复制，可扫码发送" : "复制咨询摘要"}</button></div>}
        </form>
      </div>
    </section>
  );
}
