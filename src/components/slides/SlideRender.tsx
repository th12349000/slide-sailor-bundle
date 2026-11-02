import { createContext, useContext, useEffect, useRef, useState, type ReactNode } from "react";
import { FONT_STACK, SLIDE_H, SLIDE_W, type Slide, type SlideElement } from "@/lib/deck";

const ScaleContext = createContext(1);
export const useSlideScale = () => useContext(ScaleContext);

export function ScaledSlide({ children, className = "" }: { children: ReactNode; className?: string }) {
  const ref = useRef<HTMLDivElement>(null);
  const [scale, setScale] = useState(0.1);
  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    const ro = new ResizeObserver(() => {
      setScale(Math.min(el.clientWidth / SLIDE_W, el.clientHeight / SLIDE_H));
    });
    ro.observe(el);
    return () => ro.disconnect();
  }, []);
  return (
    <div ref={ref} className={`relative overflow-hidden ${className}`} data-scale={scale}>
      <div
        style={{
          position: "absolute",
          width: SLIDE_W,
          height: SLIDE_H,
          left: "50%",
          top: "50%",
          marginLeft: -SLIDE_W / 2,
          marginTop: -SLIDE_H / 2,
          transform: `scale(${scale})`,
          transformOrigin: "center center",
        }}
      >
        <ScaleContext value={scale}>{children}</ScaleContext>
      </div>
    </div>
  );
}


const POLY: Partial<Record<SlideElement["shape"], string>> = {
  triangle: "50,0 100,100 0,100",
  diamond: "50,0 100,50 50,100 0,50",
  star: "50,0 61,35 98,35 68,57 79,91 50,70 21,91 32,57 2,35 39,35",
  arrow: "0,30 65,30 65,0 100,50 65,100 65,70 0,70",
};

export function ElementBody({ el, editingText }: { el: SlideElement; editingText?: ReactNode }) {
  const textStyle: React.CSSProperties = {
    fontSize: el.fontSize,
    fontFamily: FONT_STACK[el.fontFamily],
    color: el.color,
    fontWeight: el.bold ? 700 : 400,
    fontStyle: el.italic ? "italic" : "normal",
    textDecoration: el.underline ? "underline" : "none",
    textAlign: el.align,
    lineHeight: 1.15,
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
  };
  const text = editingText ?? (el.text ? <div style={textStyle}>{el.text}</div> : null);

  if (el.type === "image") {
    return el.src ? (
      <img
        src={el.src}
        alt=""
        draggable={false}
        style={{ width: "100%", height: "100%", objectFit: "cover", borderRadius: el.radius, border: el.borderWidth ? `${el.borderWidth}px solid ${el.borderColor}` : undefined }}
      />
    ) : (
      <div style={{ width: "100%", height: "100%", background: "#ddd", display: "grid", placeItems: "center", fontSize: 32, color: "#777" }}>No image</div>
    );
  }

  if (el.type === "shape" && (POLY[el.shape] || el.shape === "line")) {
    return (
      <div style={{ position: "relative", width: "100%", height: "100%" }}>
        <svg viewBox="0 0 100 100" preserveAspectRatio="none" style={{ position: "absolute", inset: 0, width: "100%", height: "100%", overflow: "visible" }}>
          {el.shape === "line" ? (
            <line x1="0" y1="50" x2="100" y2="50" stroke={el.fill === "transparent" ? el.borderColor : el.fill} strokeWidth={Math.max(4, el.borderWidth)} vectorEffect="non-scaling-stroke" />
          ) : (
            <polygon points={POLY[el.shape]} fill={el.fill} stroke={el.borderColor} strokeWidth={el.borderWidth} vectorEffect="non-scaling-stroke" />
          )}
        </svg>
        <div style={{ position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
          <div style={{ width: "100%" }}>{text}</div>
        </div>
      </div>
    );
  }

  const isShape = el.type === "shape";
  return (
    <div
      style={{
        width: "100%",
        height: "100%",
        background: el.fill,
        borderRadius: isShape && el.shape === "ellipse" ? "50%" : el.radius,
        border: el.borderWidth ? `${el.borderWidth}px solid ${el.borderColor}` : undefined,
        display: "flex",
        alignItems: isShape ? "center" : "flex-start",
        padding: isShape ? 32 : el.fill !== "transparent" ? 24 : 0,
        boxSizing: "border-box",
      }}
    >
      <div style={{ width: "100%" }}>{text}</div>
    </div>
  );
}

export function elementBoxStyle(el: SlideElement): React.CSSProperties {
  return {
    position: "absolute",
    left: el.x,
    top: el.y,
    width: el.w,
    height: el.h,
    transform: el.rotation ? `rotate(${el.rotation}deg)` : undefined,
    opacity: el.opacity,
  };
}

/** Static render of a slide (thumbnails, viewer). */
export function SlideStatic({ slide, onLink }: { slide: Slide; onLink?: (el: SlideElement) => void }) {
  return (
    <div style={{ width: SLIDE_W, height: SLIDE_H, background: slide.background, position: "relative", overflow: "hidden" }}>
      {slide.elements.map((el) => (
        <div
          key={el.id}
          style={{ ...elementBoxStyle(el), cursor: el.link && onLink ? "pointer" : undefined }}
          onClick={el.link && onLink ? (e) => { e.stopPropagation(); onLink(el); } : undefined}
          title={el.link?.kind === "url" ? el.link.url : undefined}
        >
          <ElementBody el={el} />
        </div>
      ))}
    </div>
  );
}
