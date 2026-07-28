
(() => {
  const modal = document.querySelector("[data-assistant-modal]");
  const openers = document.querySelectorAll("[data-assistant-open]");
  const closers = document.querySelectorAll("[data-assistant-close]");
  const messages = document.querySelector("[data-assistant-messages]");
  const answers = {
    business: "早舟聚焦三条主线：产教融合与产业学院、学校内涵建设、AI赋能教育与产教融合工作场景。可以从核心业务或案例成果页进一步了解。",
    product: "如果只是持续学习，可先了解399元成长圈；有明确单点问题可考虑专题工作坊；正在推进真实项目，可咨询旗舰共创或机构定制。最终方案需由工作人员结合单位、身份与项目阶段确认。",
    training: "2026年下半年已安排AI赋能、企业产教融合能力提升、项目操盘与成果共创等主题。请在“培训日历”查看时间，并联系朱老师预报名。",
    human: "工作坊与课程预约请联系朱老师：152 1321 7592；产业学院、学校内涵建设、AI工作流及机构项目请联系杜老师：186 9698 7413。"
  };
  function openModal() {
    if (!modal) return;
    modal.hidden = false;
    document.body.style.overflow = "hidden";
    modal.querySelector("[data-assistant-close]")?.focus();
  }
  function closeModal() {
    if (!modal) return;
    modal.hidden = true;
    document.body.style.overflow = "";
  }
  openers.forEach((button) => button.addEventListener("click", openModal));
  closers.forEach((button) => button.addEventListener("click", closeModal));
  document.addEventListener("keydown", (event) => { if (event.key === "Escape") closeModal(); });
  document.querySelectorAll("[data-question]").forEach((button) => {
    button.addEventListener("click", () => {
      if (!messages) return;
      const user = document.createElement("p");
      user.className = "user";
      user.textContent = button.textContent;
      const bot = document.createElement("p");
      bot.className = "bot";
      bot.textContent = answers[button.dataset.question] || "这个问题需要工作人员进一步确认。";
      messages.append(user, bot);
      messages.scrollTop = messages.scrollHeight;
    });
  });
})();
