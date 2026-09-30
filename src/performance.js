const ANALYZER_TIMING_LOG_PREFIX = "analyzer.performance";

function createAnalyzerPhaseTimer_(context) {
  const startedAt = Date.now();
  const phases = {};
  let finished = false;
  const details = Object.assign({}, context || {});

  return {
    start: () => Date.now(),
    end(phase, phaseStartedAt) {
      if (!phase || phaseStartedAt == null) return;
      const elapsed = Math.max(0, Date.now() - phaseStartedAt);
      phases[phase] = (phases[phase] || 0) + elapsed;
    },
    setContext(values) {
      Object.assign(details, values || {});
    },
    finish(outcome) {
      if (finished) return;
      finished = true;
      const entry = Object.assign({}, details, {
        outcome: outcome || "unknown",
        totalMs: Math.max(0, Date.now() - startedAt),
        phasesMs: phases
      });
      console.log(ANALYZER_TIMING_LOG_PREFIX + " " + JSON.stringify(entry));
    }
  };
}

function startAnalyzerPhase_(timer) {
  return timer && typeof timer.start === "function" ? timer.start() : Date.now();
}

function endAnalyzerPhase_(timer, phase, startedAt) {
  if (timer && typeof timer.end === "function") timer.end(phase, startedAt);
}
