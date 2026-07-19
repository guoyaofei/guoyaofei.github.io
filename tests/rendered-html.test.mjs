import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

async function render() {
  const workerUrl = new URL("../dist/server/index.js", import.meta.url);
  workerUrl.searchParams.set("test", `${process.pid}-${Date.now()}`);
  const { default: worker } = await import(workerUrl.href);

  return worker.fetch(
    new Request("http://localhost/", { headers: { accept: "text/html" } }),
    { ASSETS: { fetch: async () => new Response("Not found", { status: 404 }) } },
    { waitUntil() {}, passThroughOnException() {} },
  );
}

test("server-renders the zaozhou corporate homepage", async () => {
  const response = await render();
  assert.equal(response.status, 200);
  assert.match(response.headers.get("content-type") ?? "", /^text\/html\b/i);

  const html = await response.text();
  assert.match(html, /<title>早舟科技｜让产教融合从合作走向成果<\/title>/i);
  assert.match(html, /zaozhou-symbol\.png/);
  assert.match(html, /zaozhou-wordmark-v3\.png/);
  assert.match(html, /20 年职业教育实践/);
  assert.match(html, /500/);
  assert.match(html, /杜平（平哥）/);
  assert.match(html, /pingge-portrait\.jpg/);
  assert.match(html, /平哥产教融合成长圈/);
  assert.match(html, /产教融合专题实战工作坊/);
  assert.match(html, /项目操盘与成果共创旗舰工作坊/);
  assert.match(html, /产教融合项目操盘手研修计划/);
  assert.match(html, /产教融合事业大会员/);
  assert.match(html, /机构定制解决方案/);
  assert.match(html, /视频 · 文章 · 直播 · 行业洞察/);
  assert.match(html, /og\.jpg/);
  assert.match(html, /平哥随笔/);
  assert.match(html, /AI 工作流/);
  assert.match(html, /需求预诊断/);
  assert.match(html, /早舟智能咨询助手/);
  assert.match(html, /在线咨询/);
  assert.match(html, /官网智能导览/);
  assert.match(html, /更多公开报道与成果线索/);
  assert.match(html, /关注平哥，持续获取一线观点与方法/);
  assert.match(html, /wechat-pingge-talk-video\.jpg/);
  assert.match(html, /wechat-pingge-life-video\.jpg/);
  assert.match(html, /wechat-pingge-suibi\.jpg/);
  assert.match(html, /wechat-pingge-business\.jpg/);
  assert.match(html, /查看适合对象与交付结果/);
  assert.doesNotMatch(html, /codex-preview|Your site is taking shape|Codex is working/i);
});

test("keeps the local preview self-contained and the hosted page lightweight", async () => {
  const html = await readFile(new URL("../outputs/早舟科技官网-本地预览.html", import.meta.url), "utf8");
  const hosted = await readFile(new URL("../work/github-pages/index.html", import.meta.url), "utf8");
  assert.match(html, /data:image\/png;base64/);
  assert.match(html, /data:image\/jpeg;base64/);
  assert.match(html, /class="brand-wordmark"/);
  assert.match(html, /长春科技学院学校管理能力提升活动现场/);
  assert.match(html, /平哥产教融合成长圈/);
  assert.match(html, /class="proof-grid"/);
  assert.match(html, /class="journey-path"/);
  assert.match(html, /class="offer-grid"/);
  assert.match(html, /class="video-topic-grid"/);
  assert.match(html, /class="report-index"/);
  assert.match(html, /id="diagnosis"/);
  assert.match(html, /平哥聊职教生活号/);
  assert.match(html, /justify-content:\s*flex-start/);
  assert.match(hosted, /assets\/zaozhou-wordmark-v3\.png/);
  assert.match(hosted, /assets\/wechat-workshop-qr\.jpg/);
  assert.match(hosted, /assets\/og\.jpg/);
  assert.match(hosted, /平哥聊职教视频号二维码/);
  assert.match(hosted, /平哥聊职教生活号视频号二维码/);
  assert.match(hosted, /平哥随笔公众号二维码/);
  assert.match(hosted, /平哥产教融合生意经公众号二维码/);
  for (const id of ["about", "solutions", "workflow", "products", "cases", "media", "diagnosis", "contact"]) {
    assert.match(hosted, new RegExp(`id="${id}"`));
  }
  assert.match(hosted, /工作坊与课程预约/);
  assert.match(hosted, /项目与专家咨询/);
  assert.match(hosted, /WG77fME9DlXbazd62AEcNvaRnUd/);
  assert.match(hosted, /¥100,000 起/);
  assert.doesNotMatch(hosted, /职业教育 · 产教融合 · AI<\/small>/);
  assert.doesNotMatch(hosted, /二维码待提供|收到原图后替换|第5版新功能|演示版在线|扫码直达，不再让客户手动搜索/);
});
