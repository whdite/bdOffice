import { BIGDOG_REPOSITORY } from '../../shared/bigdog-release'

/** Fork-owned product copy. Other UI locales use English for these new strings. */
const EN = {
  product: 'BigDog Office',
  productDescription:
    'Edit documents, spreadsheets, presentations and PDFs locally. AI BigDog can help you review and revise files.',
  website: 'Official website',
  websiteOpenFailed:
    'Could not open the browser. You can copy the website link and open it manually.',
  credits: 'Built on GenOffice',
  attribution: 'Built on GenOffice. Thank you to the GenOffice team and open-source contributors.',
  licenses: 'Open-source licenses and third-party notices are included with the app.',
  releases: 'Software updates',
  releaseDescription:
    'Check the selected channel for a version compatible with this device. Release notes and installers are available on GitHub.',
  repository: `GitHub · ${BIGDOG_REPOSITORY}`,
  newVersion: 'Update available',
  dismissUpdate: 'Remind me next time',
  openReleases: 'View releases',
  checkUpdates: 'Check for updates',
  checking: 'Checking…',
  upToDate: 'No newer compatible release is available in this channel.',
  checkFailed: 'Could not check for updates. Please try again or visit Releases.',
  available: (version: string) =>
    `Version ${version} is available. Download the installer from Releases.`,
  slides: [
    {
      title: 'Welcome to BigDog Office',
      subtitle: 'Your local AI office companion',
      body: 'Create documents, work with spreadsheets, build presentations and review PDFs with AI by your side.',
      art: 'logo',
    },
    {
      title: 'Built on GenOffice, with thanks',
      subtitle: 'BigDog Office is built on the open-source GenOffice project.',
      body: 'Thank you to the GenOffice team and the open-source community for the editing and rendering foundations. BigDog Office adds its own workspace Agent and workflow improvements.',
      art: 'gift',
    },
    {
      title: 'Make it your workspace',
      subtitle: 'Choose your AI provider in Settings to get started.',
      body: 'Configure an API or connect an account. Availability, usage limits and charges follow your provider’s terms.',
      art: 'check',
    },
  ],
} as const

const ZH: BigDogCopy = {
  product: '大狗Office',
  productDescription: '在本机编辑文档、表格、演示文稿和 PDF，也可以请 AI 大狗整理资料、修改文件。',
  website: '官方网站',
  websiteOpenFailed: '暂时无法打开浏览器，请复制官网链接后访问。',
  credits: '基于 GenOffice 构建 · 致谢',
  attribution: '基于 GenOffice 构建，感谢原项目团队与所有开源贡献者。',
  licenses: '开源许可证和第三方组件声明随应用提供。',
  releases: '版本更新',
  releaseDescription: '按所选通道检查适用于本机的版本。更新说明和安装包可在 GitHub 发布页查看。',
  repository: `GitHub · ${BIGDOG_REPOSITORY}`,
  newVersion: '发现新版',
  dismissUpdate: '下次启动再提醒',
  openReleases: '查看发布版本',
  checkUpdates: '检查更新',
  checking: '正在检查…',
  upToDate: '当前通道暂无适用于本机的更新版本。',
  checkFailed: '暂时无法检查更新，请重试或前往发布页面。',
  available: (version) => `发现新版本 ${version}，可前往发布页面下载安装。`,
  slides: [
    {
      title: '欢迎使用大狗Office',
      subtitle: '你的本地 AI 办公伙伴',
      body: '创建文档、处理表格、制作演示与审阅 PDF，让 AI 参与你的日常办公。',
      art: 'logo',
    },
    {
      title: '源自 GenOffice，感谢开源',
      subtitle: '大狗Office 基于开源项目 GenOffice 构建。',
      body: '感谢 GenOffice 团队与开源社区提供的编辑和渲染基础。大狗Office 在此基础上拓展工作区 Agent 与办公流程。',
      art: 'gift',
    },
    {
      title: '准备好你的工作区',
      subtitle: '在设置中选择 AI 服务商，即可开始。',
      body: '配置 API 或连接账户即可使用。服务额度与费用以服务商规则为准。',
      art: 'check',
    },
  ],
}

type BigDogCopy = {
  [K in keyof typeof EN]: K extends 'slides'
    ? readonly { title: string; subtitle: string; body: string; art: 'logo' | 'gift' | 'check' }[]
    : K extends 'available'
      ? (version: string) => string
      : string
}

const ZH_TW: BigDogCopy = {
  ...ZH,
  productDescription: '在本機編輯文件、試算表、簡報與 PDF，也可以請 AI 大狗整理資料、修改檔案。',
  website: '官方網站',
  websiteOpenFailed: '暫時無法開啟瀏覽器，請複製官網連結後造訪。',
  credits: '基於 GenOffice 建構 · 致謝',
  attribution:
    '大狗Office 基於開源專案 GenOffice 建構，增加本機工作區 Agent 與辦公流程改進。感謝 GenOffice 團隊及所有開源貢獻者。',
  licenses: '開源授權與第三方元件聲明隨應用程式提供。',
  releases: '版本更新',
  releaseDescription:
    '依所選通道檢查適用於這台裝置的版本。更新說明與安裝檔可在 GitHub 發布頁面查看。',
  newVersion: '發現新版',
  dismissUpdate: '下次啟動再提醒',
  openReleases: '查看發佈版本',
  checkUpdates: '檢查更新',
  checking: '正在檢查…',
  upToDate: '目前通道暫無適用於本機的更新版本。',
  checkFailed: '暫時無法檢查更新，請重試或前往發佈頁面。',
  available: (version) => `發現新版本 ${version}，可前往發佈頁面下載安裝。`,
  slides: [
    {
      title: '歡迎使用大狗Office',
      subtitle: '你的本機 AI 辦公夥伴',
      body: '建立文件、處理試算表、製作簡報與審閱 PDF，讓 AI 參與你的日常辦公。',
      art: 'logo',
    },
    {
      title: '源自 GenOffice，感謝開源',
      subtitle: '大狗Office 基於開源專案 GenOffice 建構。',
      body: '感謝 GenOffice 團隊與開源社群提供的編輯和渲染基礎。大狗Office 在此基礎上拓展工作區 Agent 與辦公流程。',
      art: 'gift',
    },
    {
      title: '準備好你的工作區',
      subtitle: '在設定中選擇 AI 服務商，即可開始。',
      body: '設定 API 或連接帳戶即可使用。服務額度與費用以服務商規則為準。',
      art: 'check',
    },
  ],
}

export function bigDogCopy(lang: string): BigDogCopy {
  return lang === 'zh-TW' ? ZH_TW : lang === 'zh' ? ZH : EN
}
