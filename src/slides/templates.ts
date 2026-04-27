export const COLORS = {
  primary: { red: 0.2, green: 0.4, blue: 0.8 },
  success: { red: 0.2, green: 0.7, blue: 0.3 },
  warning: { red: 0.9, green: 0.6, blue: 0.1 },
  danger: { red: 0.8, green: 0.2, blue: 0.2 },
  dark: { red: 0.2, green: 0.2, blue: 0.2 },
  light: { red: 0.95, green: 0.95, blue: 0.95 },
  white: { red: 1, green: 1, blue: 1 },
};

export function getStatusColor(status: "pass" | "warning" | "issue") {
  switch (status) {
    case "pass":
      return COLORS.success;
    case "warning":
      return COLORS.warning;
    case "issue":
      return COLORS.danger;
  }
}

export function getStatusEmoji(status: "pass" | "warning" | "issue"): string {
  switch (status) {
    case "pass":
      return "OK";
    case "warning":
      return "WARN";
    case "issue":
      return "ISSUE";
  }
}

export function getRiskColor(risk: "low" | "medium" | "high") {
  switch (risk) {
    case "low":
      return COLORS.success;
    case "medium":
      return COLORS.warning;
    case "high":
      return COLORS.danger;
  }
}

export function getVerdictColor(verdict: "APPROVE" | "REQUEST_CHANGES" | "COMMENT") {
  switch (verdict) {
    case "APPROVE":
      return COLORS.success;
    case "REQUEST_CHANGES":
      return COLORS.danger;
    case "COMMENT":
      return COLORS.warning;
  }
}
