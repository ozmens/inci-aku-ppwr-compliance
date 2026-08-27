import { useEffect, useState } from "react";
import { applyTheme, getTheme, type Theme } from "../theme";

export default function ThemeToggle({
  compact = false,
  variant = "switch",
}: {
  compact?: boolean;
  variant?: "switch" | "bar";
}) {
  const [theme, setTheme] = useState<Theme>(() => getTheme());

  useEffect(() => {
    applyTheme(theme);
  }, [theme]);

  const next: Theme = theme === "dark" ? "light" : "dark";

  if (variant === "bar") {
    return (
      <button
        type="button"
        className="theme-bar-btn"
        title={next === "light" ? "Açık tema" : "Koyu tema"}
        aria-label={next === "light" ? "Açık temaya geç" : "Koyu temaya geç"}
        onClick={() => setTheme(next)}
      >
        {next === "light" ? "Açık" : "Koyu"}
      </button>
    );
  }

  return (
    <div className={`theme-switch ${compact ? "compact" : ""}`} role="group" aria-label="Tema">
      <button
        type="button"
        className={theme === "light" ? "on" : ""}
        aria-pressed={theme === "light"}
        title="Açık tema"
        onClick={() => setTheme("light")}
      >
        {compact ? "☀" : "Açık"}
      </button>
      <button
        type="button"
        className={theme === "dark" ? "on" : ""}
        aria-pressed={theme === "dark"}
        title="Koyu tema"
        onClick={() => setTheme("dark")}
      >
        {compact ? "☾" : "Koyu"}
      </button>
    </div>
  );
}
