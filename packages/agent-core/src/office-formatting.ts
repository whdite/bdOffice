/** Let an explicit formatting request continue through the format-specific tools. */
export function hasExplicitOfficeFormatting(instruction: string): boolean {
  return (
    /模板|母版|配色|色系|风格|品牌(?:色|规范)|版式|排版|页边距|页眉|页脚|行距|段[前后]|[大小一二三四五六七八初]号字|字号|字体|宋体|黑体|楷体|仿宋|微软雅黑|[左右]对齐|居中|两端对齐|列宽|行高|表头.{0,8}(填充|[红蓝绿紫橙白黑灰黄]色|加粗|对齐)|(?:背景|标题|主色).{0,8}[红蓝绿紫橙白黑灰黄]色|\d\s*[:：]\s*\d+\s*(?:比例|画幅)/i.test(
      instruction,
    ) ||
    /(?:#(?:[\da-f]{6}|[\da-f]{3})\b|\b(?:template|slide master|style|font|typeface|layout|typography|palette|colou?r scheme|line spacing|margins?|row height|column width|16\s*:\s*9|4\s*:\s*3)\b)/i.test(
      instruction,
    )
  )
}
