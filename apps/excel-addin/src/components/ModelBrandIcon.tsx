import type { ReactNode } from "react";
import deepseek from "../assets/model-brands/deepseek.svg";
import openai from "../assets/model-brands/openai.svg";
import anthropic from "../assets/model-brands/anthropic.svg";
import google from "../assets/model-brands/google.svg";
import qwen from "../assets/model-brands/qwen.svg";
import moonshot from "../assets/model-brands/moonshot.svg";
import zhipu from "../assets/model-brands/zhipu.svg";
import mistral from "../assets/model-brands/mistral.svg";
import meta from "../assets/model-brands/meta.svg";
import xai from "../assets/model-brands/xai.svg";

// Match upstream model families, including provider-prefixed IDs. Connection
// nicknames and the OpenAI-compatible transport do not identify the vendor.
const brands = [
  { pattern: /^(deepseek)(?:[-_.:]|$)/i, name: "DeepSeek", icon: deepseek },
  { pattern: /^(gpt|chatgpt|o[134])(?:[-_.:]|$)/i, name: "OpenAI", icon: openai },
  { pattern: /^claude(?:[-_.:]|$)/i, name: "Anthropic", icon: anthropic },
  { pattern: /^(gemini|gemma)(?:[-_.:]|\d|$)/i, name: "Google", icon: google },
  { pattern: /^(qwen|qwq)(?:[-_.:]|\d|$)/i, name: "Qwen", icon: qwen },
  { pattern: /^(kimi|moonshot)(?:[-_.:]|$)/i, name: "Moonshot", icon: moonshot },
  { pattern: /^(glm|chatglm)(?:[-_.:]|\d|$)/i, name: "Zhipu", icon: zhipu },
  { pattern: /^(mistral|mixtral|codestral|magistral|devstral|pixtral|ministral)(?:[-_.:]|$)/i, name: "Mistral", icon: mistral },
  { pattern: /^(llama|meta-llama)(?:[-_.:]|\d|$)/i, name: "Meta", icon: meta },
  { pattern: /^grok(?:[-_.:]|$)/i, name: "xAI", icon: xai },
];

export function ModelBrandIcon({ modelId, fallback }: { modelId: string; fallback: ReactNode }) {
  const family = modelId.split("/").at(-1) ?? "";
  const brand = brands.find(({ pattern }) => pattern.test(family));
  if (!brand) return fallback;
  // Bundlers can inline SVG as a data URI containing literal quotes.
  const mask = `url("${brand.icon.replace(/["\\\r\n\f]/g, (character) => encodeURIComponent(character))}")`;
  return <i className="model-picker-symbol model-provider-logo" data-model-brand={brand.name}
    aria-hidden="true" title={brand.name}
    style={{ maskImage: mask, WebkitMaskImage: mask }} />;
}
