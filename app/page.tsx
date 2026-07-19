import Image from "next/image";
import ConsultationForm from "./ConsultationForm";
import SmartAssistant from "./SmartAssistant";

const navItems = [
  ["关于早舟", "#about"],
  ["解决方案", "#solutions"],
  ["产品服务", "#products"],
  ["案例成果", "#cases"],
  ["内容中心", "#media"],
  ["联系我们", "#contact"],
] as const;

const services = [
  {
    number: "01",
    title: "产教融合与产业学院",
    text: "从合作模式设计、专业群与产业链对接，到产业学院实体化运营、知识资产化与成果评价，让校企合作真正产生价值。",
    points: ["产业学院规划与运营", "产教融合项目策划", "校企合作机制建设"],
  },
  {
    number: "02",
    title: "学校内涵建设",
    text: "围绕党建、战略、教学、德育与文化建设，提供从规划到成果表达的全过程专业支持。",
    points: ["规划·计划·课题研究", "成果打造·亮点提炼", "教师培训·教材编制"],
  },
  {
    number: "03",
    title: "AI 赋能教育场景",
    text: "把人工智能嵌入教育教学与产教融合工作流，帮助学校和企业提升调研、方案、运营与成果沉淀效率。",
    points: ["AI 教学应用", "AI 产教融合工作流", "组织能力与教师培训"],
  },
];

const workflow = [
  { step: "01", title: "需求诊断", text: "访谈、数据与政策协同分析，识别真正需要解决的问题。" },
  { step: "02", title: "智能策划", text: "专家经验叠加 AI，快速形成可论证、可执行的建设方案。" },
  { step: "03", title: "实体运营", text: "将方案落到组织、项目、课程、企业资源与日常工作机制。" },
  { step: "04", title: "成果量化", text: "沉淀案例、标准、教材、课题与数据资产，形成可传播成果。" },
];

const productJourney = [
  {
    stage: "认识",
    price: "免费",
    title: "职业教育与产教融合内容矩阵",
    result: "看见平哥，理解趋势与行动机会",
    text: "通过视频号直播、短视频、公众号、公开课和朋友圈，持续提供政策解读、行业判断、案例分析与实战观点。",
    delivery: "直播 · 短视频 · 公众号 · 公开课",
    audience: "适合：第一次认识平哥、希望持续获取行业判断的人",
    gain: "可获得：政策解读、案例观点和进入后续产品的清晰入口",
  },
  {
    stage: "持续成长",
    price: "¥399 / 年",
    title: "平哥产教融合成长圈",
    result: "看懂趋势、发现问题、明确下一步",
    text: "全年约40场线上成长夜话，持续解读政策、案例、项目与方法，并提供公开项目会诊、工具资料和专题工作坊优先报名资格。",
    delivery: "线上直播 · 案例会诊 · 社群运营 · 资料分享",
    featured: true,
    audience: "适合：院校干部、教师、企业项目人员及职业教育同行",
    gain: "可获得：全年约40场直播、案例会诊、工具资料与社群连接",
  },
  {
    stage: "单点突破",
    price: "¥8,800",
    title: "产教融合专题实战工作坊",
    result: "两天集中解决一个具体问题",
    text: "设置AI赋能产教融合、企业进入职业院校、职教专家IP与知识产品化、产业学院运营与项目设计等专题。",
    delivery: "两天一夜 · 小班共创 · 工具模板 · 成果点评",
    audience: "适合：已有明确单点问题、希望两天集中突破的客户",
    gain: "可获得：一个专题成果、工具模板、行动清单与学员社群",
  },
  {
    stage: "项目共创",
    price: "¥19,800",
    title: "项目操盘与成果共创旗舰工作坊",
    result: "带着真实问题来，带着项目方案走",
    text: "从课前诊断、机制设计到成果路径和90天行动清单，帮助客户完成复杂项目从判断到落地的闭环。",
    delivery: "三天两夜 · 课前诊断 · 一对一咨询 · 阶段跟踪",
    audience: "适合：正在承担真实产教融合或产业学院项目的负责人",
    gain: "可获得：诊断报告、机制设计、成果路径与90天行动清单",
  },
  {
    stage: "能力升级",
    price: "¥39,800",
    title: "产教融合项目操盘手研修计划",
    result: "从听懂产教融合到能够独立操盘",
    text: "面向高潜项目负责人，以3至6个月周期培养战略判断、方案写作、商务谈判、组织协调和成果交付能力。",
    delivery: "申请制 · 真实项目 · 导师辅导 · 路演答辩",
    audience: "适合：希望成为产业学院或产教融合项目操盘者的人",
    gain: "可获得：3—6个月训练、真实项目、导师辅导和个人操盘案例",
  },
  {
    stage: "事业协同",
    price: "¥100,000 / 年",
    title: "平哥产教融合事业大会员",
    result: "获得长期战略、项目与事业支持",
    text: "面向高信任学校管理者、企业负责人和项目操盘者，提供年度诊断、固定次数咨询、重大项目会诊与闭门共创。",
    delivery: "邀请审核制 · 年度顾问 · 项目会诊 · 资源协同",
    audience: "适合：学校管理者、企业负责人和高端项目操盘者",
    gain: "可获得：年度战略诊断、重大项目会诊、闭门共创与事业复盘",
  },
  {
    stage: "组织成果",
    price: "¥100,000 起",
    title: "产教融合机构定制解决方案",
    result: "建成真实项目，形成组织标杆",
    text: "面向院校、央国企、产业园区和行业组织，提供顶层设计、产业学院建设运营、专业群设计、项目陪跑和成果提炼。",
    delivery: "调研诊断 · 方案共创 · 项目推进 · 结项复盘",
    audience: "适合：院校、二级学院、央国企、产业园区和行业组织",
    gain: "可获得：从调研到结项的项目制服务与可展示的组织成果",
  },
];

const wechatAccounts = [
  { type: "视频号", name: "平哥聊职教", note: "政策解读 · 产教融合 · 职教观察", qr: "/brand/wechat-pingge-talk-video.jpg" },
  { type: "视频号", name: "平哥聊职教生活号", note: "同行交流 · 创业与职业教育实践", qr: "/brand/wechat-pingge-life-video.jpg" },
  { type: "公众号", name: "平哥随笔", note: "职业教育与产教融合原创文章", qr: "/brand/wechat-pingge-suibi.jpg" },
  { type: "公众号", name: "平哥产教融合生意经", note: "产教融合项目与专业运营方法", qr: "/brand/wechat-pingge-business.jpg" },
] as const;

const moreReports = [
  { category: "产教融合", title: "基于产教融合的“五金建设”专题讲座", source: "南充科技职业学院", href: "https://www.nstc.edu.cn/Main/Index/detail/id/2399.html" },
  { category: "AI赋能", title: "AI如何赋能产教融合：专题研修班提供路径", source: "广州市技师学院", href: "https://www.gzgj.net/hzjl/zjjt/content_4721" },
  { category: "学校内涵", title: "“四化育三力”全程职业生涯教育改革", source: "中国职业教育网 / 中国教育报", href: "https://xhgz.chinazy.org/platform/service/zxnews/shtml/201804/13495.shtml" },
  { category: "课题研究", title: "职业教育服务乡村振兴与产教融合课题研讨", source: "重庆统一战线", href: "https://www.cqtzb.gov.cn/web/article/1415840243113201664/web/content_1415840243113201664.html" },
  { category: "现代学徒制", title: "电子商务现代学徒制项目企业实战班", source: "重庆市万盛经开区管委会", href: "https://ws.cq.gov.cn/zwgk_165/sydwfrnb/202102/t20210222_8916171.html" },
  { category: "校企合作", title: "电子商务真实项目进课堂与工学交替", source: "重庆工商学校", href: "https://www.cqics.cn/departs.html" },
] as const;

const cases = [
  {
    image: "/cases/nanchong-classroom.jpg",
    alt: "杜平在产教融合工作坊授课",
    tag: "产教融合实训",
    title: "平哥产教融合线下工作坊",
    text: "围绕产业学院运营、知识资产化与 AI 工具应用，以案例拆解和任务共创推动方案落地。",
    href: "https://www.nstc.edu.cn/Main/Index/detail/id/2428.html",
    source: "南充科技职业学院公开报道",
  },
  {
    image: "/cases/nanchong-discussion.jpg",
    alt: "工作坊学员参与讨论",
    tag: "产业学院诊断",
    title: "产业学院实体化运营辅导",
    text: "面向养老护理、无人机、跨境电商、汽车维修等方向，研讨专业与产业资源的深度连接。",
    href: "https://www.ncjsxy.net/Index/detail/id/1579.html",
    source: "南充技师学院公开报道",
  },
  {
    image: "/cases/changchun-leadership.jpg",
    alt: "长春科技学院学校管理能力提升活动现场",
    tag: "AI 赋能管理",
    title: "AI 时代学校管理能力提升",
    text: "为院校管理者构建战略思维、创新方法与数字化工具相结合的行动框架。",
    href: "https://www.cstu.edu.cn/dzbgs/info/1005/2152.htm",
    source: "长春科技学院公开报道",
  },
];

const milestones = [
  ["2007", "九源职业培训学校成立，从技能人才培养和校企合作起步。"],
  ["2011", "由学生培训转向教师培训，推出教师企业意识体验培训。"],
  ["2014", "开启定制培训，并参与获得国家级教学成果二等奖。"],
  ["2017", "注册九源教育咨询公司，由培训服务升级为改革咨询。"],
  ["2024", "成立早舟科技，原九源教育业务全面转入新的发展阶段。"],
  ["2025", "聚焦产教融合，形成工作坊、咨询与项目操盘产品体系。"],
  ["2026", "与西部产教融合研究院战略协同，强化研究与成果转化。"],
] as const;

const videoTopics = [
  ["政策与趋势", "职业教育政策解读、行业判断与“十五五”机会"],
  ["产业学院", "产业学院建设、运营机制与真实项目复盘"],
  ["AI 赋能", "AI进入教学、管理与产教融合工作流的方法"],
  ["项目操盘", "企业入校、方案设计、成果共创与专家IP"],
] as const;

const articles = [
  {
    type: "公开报道",
    title: "从方案到利润：产教融合 3.0 知识资产化系统实训",
    href: "https://www.nstc.edu.cn/Main/Index/detail/id/2428.html",
  },
  {
    type: "研修动态",
    title: "AI 赋能产教融合专题研修：让技术真正进入工作场景",
    href: "https://www.gzgj.net/hzjl/zjjt/content_4721",
  },
  {
    type: "平哥随笔",
    title: "产教融合不是项目清单，而是可持续运行的价值系统",
    href: "https://www.nchvc.edu.cn/ggwhb/info/1009/7353.htm",
  },
];

const organizationSchema = {
  "@context": "https://schema.org",
  "@type": "Organization",
  name: "重庆市早舟科技有限公司",
  alternateName: "重庆市九源教育",
  description: "面向全国院校与产教融合企业，提供产业学院、学校内涵建设和人工智能赋能教育场景服务。",
  founder: { "@type": "Person", name: "杜平" },
  contactPoint: [
    { "@type": "ContactPoint", telephone: "+86-152-1321-7592", contactType: "客户咨询" },
    { "@type": "ContactPoint", telephone: "+86-186-9698-7413", contactType: "项目咨询" },
  ],
};

export default function Home() {
  return (
    <>
      <a className="skip-link" href="#main">跳到主要内容</a>
      <header className="site-header">
        <div className="header-inner">
          <a className="brand" href="#top" aria-label="早舟科技首页">
            <span className="brand-mark" aria-hidden="true">
              <Image src="/brand/zaozhou-symbol.png" alt="" width={512} height={512} priority />
            </span>
            <span className="brand-copy">
              <span className="brand-wordmark">
                <Image src="/brand/zaozhou-wordmark-v3.png" alt="早舟®" width={755} height={420} priority />
              </span>
            </span>
          </a>
          <nav className="desktop-nav" aria-label="主导航">
            {navItems.map(([label, href]) => <a key={href} href={href}>{label}</a>)}
          </nav>
          <SmartAssistant />
          <a className="header-cta" href="#diagnosis">预约诊断</a>
          <details className="mobile-menu">
            <summary aria-label="打开导航">目录</summary>
            <nav aria-label="移动端导航">
              {navItems.map(([label, href]) => <a key={href} href={href}>{label}</a>)}
              <a href="#diagnosis">预约诊断</a>
            </nav>
          </details>
        </div>
      </header>

      <main id="main">
        <section className="hero" id="top">
          <div className="hero-rule" aria-hidden="true" />
          <div className="hero-content">
            <p className="eyebrow">20 年职业教育实践 · 500+ 院校及合作单位</p>
            <h1>让产教融合<br /><span>从合作走向成果</span></h1>
            <p className="hero-lead">
              早舟科技以专家经验、产业资源与人工智能为引擎，帮助院校和企业把复杂的改革任务，转化为可落地、可运营、可评价、可传播的真实成果。
            </p>
            <div className="hero-actions">
              <a className="button button-primary" href="#diagnosis">填写需求，快速匹配 <span>→</span></a>
              <a className="text-link" href="#solutions">查看解决方案 <span>↘</span></a>
            </div>
            <div className="hero-stats" aria-label="早舟科技关键数据">
              <div><strong>20<small>年</small></strong><span>职业教育实战积累</span></div>
              <div><strong>500<small>+</small></strong><span>院校及合作单位</span></div>
              <div><strong>20<small>万+</small></strong><span>职业教育同行关注</span></div>
            </div>
          </div>
          <div className="hero-aside" aria-label="核心能力">
            <p>ZAOZHOU / 2026</p>
            <div className="hero-card">
              <span>我们相信</span>
              <blockquote>真正的产教融合，不止于签约，更要形成持续运行的价值系统。</blockquote>
              <p>产业学院 · 内涵建设 · AI 工作流</p>
            </div>
          </div>
        </section>

        <section className="trust-strip" aria-label="服务对象与合作网络">
          <span>服务全国职业院校与本科院校</span>
          <i aria-hidden="true" />
          <span>链接产教融合国企·央企·优质民企</span>
          <i aria-hidden="true" />
          <span>深度合作：重庆市西部产教融合研究院</span>
        </section>

        <section className="section about-section" id="about">
          <div className="section-heading centered-heading">
            <p className="section-index">01 / 关于早舟</p>
            <h2>二十年，只做一件事：<br />让教育改革产生看得见的成果</h2>
            <p>重庆市早舟科技有限公司，前身为九源教育，源起于2007年成立的重庆市南岸区九源职业培训学校。团队长期深耕职业教育改革、师资培养、学校内涵建设与产教融合，坚持以产业需求驱动改革，把方案变成机制、项目与成果。</p>
          </div>
          <div className="founder-panel">
            <div className="founder-visual">
              <Image src="/brand/pingge-portrait.jpg" alt="重庆市早舟科技有限公司创始人杜平（平哥）" fill sizes="(max-width: 900px) 100vw, 46vw" priority />
              <div className="founder-caption"><span>FOUNDER</span><strong>杜平（平哥）</strong></div>
            </div>
            <div className="founder-copy">
              <p className="orange-kicker">创始人兼董事长</p>
              <h3>把近二十年一线经验，转化为学校与企业的行动方法</h3>
              <p>杜平，业内称“平哥”，职业教育与产教融合实战专家、“平哥聊职教”内容品牌创始人，参与获得国家级教学成果二等奖。长期服务职业院校决策者，并担任多家国企、央企及民营企业的产教融合顾问。</p>
              <p>他带领团队持续深入学校和企业现场，关注产业学院如何运营、学校如何形成内生能力、AI如何进入真实工作流，现任重庆市西部产教融合研究院研究员、首席顾问。</p>
              <a className="text-link" href="#media">查看平哥观点与视频 <span>→</span></a>
            </div>
          </div>
          <div className="timeline" aria-label="早舟发展脉络">
            {milestones.map(([year, text]) => (
              <div className="timeline-item" key={year}>
                <h3>{year}</h3>
                <p>{text}</p>
              </div>
            ))}
          </div>
          <div className="proof-grid" aria-label="早舟发展成果">
            <div><strong>5万+</strong><span>技能人才培养</span></div>
            <div><strong>1.5万</strong><span>教师与管理干部培训</span></div>
            <div><strong>35期</strong><span>国家及省市师资培训项目</span></div>
            <div><strong>200+</strong><span>教育咨询策划项目</span></div>
            <div><strong>20万+</strong><span>全网职业教育同行关注</span></div>
          </div>
        </section>

        <section className="section solutions-section" id="solutions">
          <div className="section-heading split-heading">
            <div>
              <p className="section-index">02 / 三大业务</p>
              <h2>从顶层设计，<br />一直走到一线成果</h2>
            </div>
            <p>为不同发展阶段的学校与企业，组合研究、咨询、培训、策划、运营和成果建设能力。</p>
          </div>
          <div className="service-grid">
            {services.map((service) => (
              <article className="service-card" key={service.number}>
                <span className="service-number">{service.number}</span>
                <h3>{service.title}</h3>
                <p>{service.text}</p>
                <ul>{service.points.map((point) => <li key={point}>{point}</li>)}</ul>
                <a href="#diagnosis" aria-label={`咨询${service.title}`}>先做需求匹配 <span>→</span></a>
              </article>
            ))}
          </div>
        </section>

        <section className="workflow-section" id="workflow">
          <div className="workflow-inner">
            <div className="workflow-heading">
              <p className="section-index light">03 / AI 工作流</p>
              <h2>专家判断 × AI 效率<br />让每一步都可追踪、可复用</h2>
              <p>AI 不是额外增加的一套工具，而是进入咨询、教学、运营与成果生产的完整工作流。</p>
            </div>
            <div className="workflow-grid">
              {workflow.map((item, index) => (
                <article className="workflow-card" key={item.step}>
                  <div className="workflow-top">
                    <span>{item.step}</span>
                    {index < workflow.length - 1 && <i aria-hidden="true">→</i>}
                  </div>
                  <h3>{item.title}</h3>
                  <p>{item.text}</p>
                </article>
              ))}
            </div>
            <a className="button button-cream" href="#diagnosis">定制您的 AI 工作流 <span>→</span></a>
          </div>
        </section>

        <section className="section products-section" id="products">
          <div className="section-heading centered-heading narrow-heading">
            <p className="section-index">04 / 产教融合产品矩阵</p>
            <h2>从被看见，到共同做成项目</h2>
            <p>不是简单划分价格，而是承接客户从建立认知、解决问题，到能力升级、项目合作和事业协同的完整成长过程。</p>
          </div>
          <div className="journey-path" aria-label="客户成长路径">
            {productJourney.map((product, index) => (
              <a href={`#product-${index + 1}`} key={product.title}>
                <span>{String(index + 1).padStart(2, "0")}</span>
                <strong>{product.stage}</strong>
              </a>
            ))}
          </div>
          <div className="offer-grid">
            {productJourney.map((product, index) => (
              <article className={`offer-card${product.featured ? " offer-card-featured" : ""}`} id={`product-${index + 1}`} key={product.title}>
                <div className="offer-card-top"><span>{String(index + 1).padStart(2, "0")} / {product.stage}</span><strong>{product.price}</strong></div>
                <h3>{product.title}</h3>
                <p className="offer-result">{product.result}</p>
                <p>{product.text}</p>
                <small>{product.delivery}</small>
                <details className="offer-details">
                  <summary>查看适合对象与交付结果</summary>
                  <p>{product.audience}</p>
                  <p>{product.gain}</p>
                </details>
                <a href="#diagnosis">判断是否适合我 <span>↗</span></a>
              </article>
            ))}
          </div>
          <p className="purchase-note">以上为公开参考产品体系；具体名额、权益、价格、合同、发票与履约规则以最终服务说明为准。事业大会员采用邀请审核制，机构项目根据复杂程度单独报价。</p>
        </section>

        <section className="section cases-section" id="cases">
          <div className="section-heading split-heading">
            <div>
              <p className="section-index">05 / 案例成果</p>
              <h2>在真实现场，<br />解决真实问题</h2>
            </div>
            <p>以下内容链接至院校官网公开报道。更多二十年案例将根据客户授权持续上线。</p>
          </div>
          <div className="case-grid">
            {cases.map((item) => (
              <article className="case-card" key={item.title}>
                <a className="case-image" href={item.href} target="_blank" rel="noreferrer">
                  <Image src={item.image} alt={item.alt} fill sizes="(max-width: 800px) 100vw, 33vw" />
                </a>
                <div className="case-copy">
                  <span>{item.tag}</span>
                  <h3>{item.title}</h3>
                  <p>{item.text}</p>
                  <a href={item.href} target="_blank" rel="noreferrer">{item.source} <span>↗</span></a>
                </div>
              </article>
            ))}
          </div>
          <div className="report-index">
            <div className="report-index-heading"><span>MORE EVIDENCE</span><h3>更多公开报道与成果线索</h3><p>按产教融合、AI赋能、学校内涵、课题研究和校企合作分类整理，点击可查看原发布单位页面。</p></div>
            <div className="report-links">
              {moreReports.map((report, index) => (
                <a href={report.href} target="_blank" rel="noreferrer" key={report.title}>
                  <span>{String(index + 1).padStart(2, "0")}</span><div><small>{report.category} · {report.source}</small><strong>{report.title}</strong></div><b>↗</b>
                </a>
              ))}
            </div>
          </div>
        </section>

        <section className="media-section" id="media">
          <div className="section media-inner">
            <div className="media-feature">
              <p className="section-index light">06 / 内容中心</p>
              <span className="video-label">视频 · 文章 · 直播 · 行业洞察</span>
              <h2>看见行业趋势，<br />找到行动方法</h2>
              <p>围绕职业教育政策、产业学院建设、AI赋能和项目操盘，持续分享一线观察、真实案例与可复用的方法。</p>
              <a className="button button-primary" href="#wechat-matrix">关注微信内容矩阵 <span>↓</span></a>
              <a className="media-secondary-link" href="https://zusfstjyj0.feishu.cn/drive/folder/WG77fME9DlXbazd62AEcNvaRnUd?from=from_copylink" target="_blank" rel="noreferrer">浏览专题资料库 ↗</a>
            </div>
            <div className="article-list">
              {articles.map((article, index) => (
                <a href={article.href} target="_blank" rel="noreferrer" key={article.title}>
                  <span>0{index + 1}</span>
                  <div><small>{article.type}</small><h3>{article.title}</h3></div>
                  <b aria-hidden="true">↗</b>
                </a>
              ))}
            </div>
            <div className="video-topic-grid" aria-label="视频主题导航">
              {videoTopics.map(([title, text], index) => (
                <a href="#wechat-matrix" key={title}>
                  <span>0{index + 1}</span><strong>{title}</strong><p>{text}</p><b>↓</b>
                </a>
              ))}
            </div>
            <div className="wechat-directory" id="wechat-matrix" aria-label="平哥官方微信内容矩阵">
              <div className="wechat-directory-heading">
                <Image src="/brand/pingge-portrait.jpg" alt="杜平（平哥）" width={112} height={112} />
                <div><span>微信内容矩阵</span><h3>关注平哥，持续获取一线观点与方法</h3></div>
              </div>
              <div className="wechat-account-grid">
                {wechatAccounts.map((account) => (
                  <article key={account.name}>
                    <div className="wechat-account-qr"><Image src={account.qr} alt={`${account.name}${account.type}二维码`} fill sizes="(max-width: 720px) 86vw, (max-width: 1100px) 40vw, 270px" /></div>
                    <small>{account.type}</small>
                    <strong>{account.name}</strong>
                    <p>{account.note}</p>
                  </article>
                ))}
              </div>
            </div>
          </div>
        </section>

        <ConsultationForm />

        <section className="contact-section" id="contact">
          <div className="section contact-inner">
            <div className="contact-heading">
              <p className="section-index">07 / 联系我们</p>
              <h2>把您的问题，<br />变成下一步行动</h2>
              <p>无论您正在筹建产业学院、推动学校内涵建设，还是希望用 AI 提升团队效率，都可以先从一次项目诊断开始。</p>
              <div className="phone-list">
                <a href="tel:15213217592"><small>朱老师 · 工作坊与服务预约</small><strong>152 1321 7592</strong></a>
                <a href="tel:18696987413"><small>杜老师 · 项目与专家咨询</small><strong>186 9698 7413</strong></a>
              </div>
            </div>
            <div className="qr-grid">
              <figure>
                <div><Image src="/brand/wechat-workshop-qr.jpg" alt="工作坊预约报名二维码" width={640} height={640} /></div>
                <figcaption><strong>工作坊与课程预约</strong><span>扫码添加朱老师，发送“姓名＋单位＋意向产品”，快速确认名额、时间与报名流程。</span></figcaption>
              </figure>
              <figure>
                <div><Image src="/brand/wechat-consulting-qr.jpg" alt="项目咨询预约报名二维码" width={640} height={640} /></div>
                <figcaption><strong>项目与专家咨询</strong><span>适合产业学院、内涵建设、AI工作流及机构合作，扫码后发送已生成的咨询摘要。</span></figcaption>
              </figure>
            </div>
          </div>
        </section>
      </main>

      <footer>
        <div className="footer-inner">
          <div><strong>重庆市早舟科技有限公司</strong><span>前身：重庆市九源教育</span></div>
          <p>产教融合 · 学校内涵建设 · AI 赋能教育</p>
          <a href="#top">返回顶部 ↑</a>
        </div>
      </footer>
      <script type="application/ld+json" dangerouslySetInnerHTML={{ __html: JSON.stringify(organizationSchema) }} />
    </>
  );
}
