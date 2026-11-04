export type TransitionType =
  | "none"
  | "cut"
  | "fade"
  | "push"
  | "wipe"
  | "zoom"
  | "flip"
  | "dissolve"
  | "split"
  | "wheel"
  | "cover"
  | "pull";

export type TransitionDirection = "l" | "r" | "u" | "d" | "in" | "out";

export type TransitionSpeed = "fast" | "medium" | "slow";

export interface PptxTransition {
  type: TransitionType;
  direction?: TransitionDirection;
  durationMs: number;
  advanceOnClick?: boolean;
}

export const DEFAULT_TRANSITION: PptxTransition = {
  type: "fade",
  durationMs: 400,
  advanceOnClick: true,
};

export const TRANSITION_SPEED_MS: Record<TransitionSpeed, number> = {
  fast: 250,
  medium: 500,
  slow: 800,
};

export const SUPPORTED_TRANSITION_OPTIONS: Array<{ id: TransitionType; label: string }> = [
  { id: "fade", label: "Fade" },
  { id: "push", label: "Push / Slide" },
  { id: "wipe", label: "Wipe" },
  { id: "zoom", label: "Zoom" },
  { id: "flip", label: "3D Flip" },
  { id: "dissolve", label: "Dissolve" },
  { id: "split", label: "Split" },
  { id: "cut", label: "Cut (Instant)" },
];
