// animationTransform.ts

import PptxGenJS from "pptxgenjs";

export interface AnimationConfig {
  type: string;
  duration?: number;
  delay?: number;
}

// 映射常见 CSS 动画名到 pptxgenjs 动画类型
const ANIMATION_MAP: Record<string, string> = {
  fadeIn: "fadeIn",
  fadeInUp: "flyInBottom",
  fadeInDown: "flyInTop",
  fadeInLeft: "flyInLeft",
  fadeInRight: "flyInRight",
  zoomIn: "zoomIn",
  bounceIn: "bounce",
  // 常见 slide/animate.css 命名
  slideInUp: "flyInBottom",
  slideInDown: "flyInTop",
  slideInLeft: "flyInLeft",
  slideInRight: "flyInRight",
};

/**
 * 获取元素动画配置
 * 优先 data-pptx-animation，其次 CSS animationName 映射
 */
export function getElementAnimation(element: Element): AnimationConfig | undefined {
  const style = window.getComputedStyle(element);

  // 显式 data 属性优先
  const dataAnim = element.getAttribute("data-pptx-animation");
  if (dataAnim) {
    const duration = parseInt(element.getAttribute("data-pptx-duration") || "1000", 10) / 1000;
    const delay = parseInt(element.getAttribute("data-pptx-delay") || "0", 10) / 1000;
    return { type: dataAnim, duration, delay };
  }

  // CSS 动画解析（取第一个）
  const animName = style.animationName;
  if (animName && animName !== "none") {
    const primaryAnim = animName.split(",")[0].trim();

    let pptAnimType = ANIMATION_MAP[primaryAnim];

    // 模糊匹配，如 animate__fadeIn
    if (!pptAnimType) {
      const key = Object.keys(ANIMATION_MAP).find((k) => primaryAnim.toLowerCase().includes(k.toLowerCase()));
      if (key) pptAnimType = ANIMATION_MAP[key];
    }

    if (pptAnimType) {
      const durationStr = style.animationDuration.split(",")[0] ?? "1s";
      const delayStr = style.animationDelay.split(",")[0] ?? "0s";
      const parseTime = (t: string) => {
        const val = parseFloat(t);
        return t.includes("ms") ? val / 1000 : val;
      };

      return {
        type: pptAnimType,
        duration: parseTime(durationStr) || 1,
        delay: parseTime(delayStr) || 0,
      };
    }
  }

  return undefined;
}
