/**
 * 品牌 logo 的**唯一来源**。两端（Excel 任务窗格 / WPS 面板）都从这里取，
 * 避免各自记一份路径、各自漂移。
 *
 * 为什么用 `import` 而不是 CSS 里的 `url()`：
 * 两端部署基路径不同 —— Excel 是默认 `base: "/"`，WPS 是 `base: "./"`。
 * CSS 里写死绝对路径（`/assets/...`）在 WPS 的子路径部署下会 404；写相对路径
 * 又会被打包到 `dist/assets/` 后按样式表位置解析成 `dist/assets/assets/...`。
 * 交给打包器解析（`import`）才能同时对两种 base 都给出正确 URL。
 *
 * 尺寸取 128：界面里的徽标最大 46px，128 在 2x/2.5x 屏上仍然锐利，且只有
 * 一份资源被打进包（Vite 会把同一个 import 复用成同一个文件）。
 */
import logo from "./assets/excel-bro-logo-128.png";

export const brandLogoUrl: string = logo;
