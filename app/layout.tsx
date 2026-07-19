import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  metadataBase: new URL("https://guoyaofei.github.io/"),
  title: {
    default: "早舟科技｜让产教融合从合作走向成果",
    template: "%s｜早舟科技",
  },
  description: "重庆市早舟科技有限公司，深耕职业教育20年，为全国院校与企业提供产业学院、学校内涵建设和AI赋能教育场景服务，并提供智能咨询与需求预诊断。",
  keywords: ["早舟科技", "杜平", "平哥", "产教融合", "产业学院", "学校内涵建设", "职业教育", "AI教育"],
  other: { "zaozhou-version": "6.0.1" },
  openGraph: {
    title: "早舟科技｜让产教融合从合作走向成果",
    description: "20年职业教育实践，为院校和企业提供产业学院、内涵建设与AI工作流解决方案。",
    type: "website",
    locale: "zh_CN",
    images: [{ url: "/og.jpg", width: 1200, height: 630, alt: "早舟科技，让产教融合从合作走向成果" }],
  },
  twitter: {
    card: "summary_large_image",
    title: "早舟科技｜让产教融合从合作走向成果",
    description: "20年职业教育实践，为院校和企业提供产业学院、内涵建设与AI工作流解决方案。",
    images: ["/og.jpg"],
  },
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}
