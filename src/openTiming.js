const DEBUG_TIMINGS = process.env.MBOX_VIEWER_DEBUG_TIMINGS === "1";

function createOpenTiming(filePath) {
  const startedAt = Date.now();
  const marks = [];

  return {
    enabled: DEBUG_TIMINGS,
    mark(label, extra = null) {
      if (!DEBUG_TIMINGS) {
        return;
      }

      marks.push({
        label,
        atMs: Date.now() - startedAt,
        ...(extra && typeof extra === "object" ? extra : {})
      });
    },
    snapshot(extra = null) {
      if (!DEBUG_TIMINGS) {
        return null;
      }

      return {
        filePath,
        totalMs: Date.now() - startedAt,
        marks: marks.slice(),
        ...(extra && typeof extra === "object" ? extra : {})
      };
    }
  };
}

module.exports = {
  DEBUG_TIMINGS,
  createOpenTiming
};
