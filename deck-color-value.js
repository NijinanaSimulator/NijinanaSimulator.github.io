(function (root) {
  "use strict";

  const COLORS = Object.freeze(["R", "O", "Y", "G", "B", "I", "V"]);
  const STATE_COUNT = 1 << COLORS.length;
  const EPSILON = 1e-12;
  const colorsForMask = (mask) => COLORS.filter((color, index) => mask & (1 << index));

  function compareMasks(left, right) {
    const leftColors = colorsForMask(left);
    const rightColors = colorsForMask(right);
    for (let index = 0; index < Math.min(leftColors.length, rightColors.length); index += 1) {
      const difference = COLORS.indexOf(leftColors[index]) - COLORS.indexOf(rightColors[index]);
      if (difference) return difference;
    }
    return leftColors.length - rightColors.length;
  }

  function overlapSize(left, right) {
    let mask = left & right;
    let size = 0;
    while (mask) {
      mask &= mask - 1;
      size += 1;
    }
    return size;
  }

  function analyze(entries) {
    const colorCounts = Object.fromEntries(COLORS.map((color) => [color, 0]));
    const actorCounts = new Map();
    let liverCount = 0;

    for (const entry of Array.isArray(entries) ? entries : []) {
      if (!entry) continue;
      const count = Number(entry.count);
      if (!Number.isFinite(count) || count <= 0) continue;

      // The caller supplies liver cards only, including any zero-color livers.
      liverCount += count;
      const source = typeof entry.colors === "string"
        ? Array.from(entry.colors)
        : Array.isArray(entry.colors) ? entry.colors : [];
      const colors = new Set(source.map((color) => String(color).trim().toUpperCase()));
      let actorMask = 0;
      for (let index = 0; index < COLORS.length; index += 1) {
        const color = COLORS[index];
        if (colors.has(color)) {
          colorCounts[color] += count;
          actorMask |= 1 << index;
        }
      }
      actorCounts.set(actorMask, (actorCounts.get(actorMask) || 0) + count);
    }

    const colorRates = Object.fromEntries(COLORS.map((color) => [
      color,
      liverCount > 0 ? colorCounts[color] / liverCount : 0,
    ]));
    const actors = [...actorCounts.entries()]
      .map(([mask, count]) => ({ mask, count }))
      .sort((left, right) => compareMasks(left.mask, right.mask));
    const liverKinds = actors.length;
    const callRoleCount = actors.reduce((sum, actor) => sum + (actor.mask ? actor.count : 0), 0);
    const missingColors = COLORS.filter((color) => colorCounts[color] === 0);
    const metadata = { liverCount, colorCounts, colorRates, liverKinds, callRoleCount, missingColors };
    if (liverCount === 0) return { ...metadata, status: "empty", points: [], states: [] };
    if (missingColors.length > 0) {
      return { ...metadata, status: "incomplete", points: [], states: [] };
    }

    // Fixed policy: types remain renewable and copy counts set their relative
    // sampling weights; this does not model actual hands, board limits, or costs.
    // For hidden U, sample only types with a color in U, weighted by their counts.
    // Zero-hit types, including colorless livers, do not probe hidden cards.
    // For a known target c, choose a type containing c: capture costs one call.
    // T(U) counts calls to clear all colors in U, which start unrevealed.
    // A call with C reveals a uniform c and costs
    // 1 + mean[T(U\c) + 1(c not in C)], averaged over the eligible type sampler.
    const states = [{
      mask: 0,
      colors: [],
      remaining: 0,
      expectedCalls: 0,
      afterReveal: 0,
      value: 0,
      eligibleCount: 0,
      averageHitRate: 0,
    }];
    for (let mask = 1; mask < STATE_COUNT; mask += 1) {
      const colors = colorsForMask(mask);
      const remaining = colors.length;
      let nextCalls = 0;
      for (let index = 0; index < COLORS.length; index += 1) {
        if (mask & (1 << index)) nextCalls += states[mask & ~(1 << index)].expectedCalls;
      }

      let eligibleCount = 0;
      let weightedCalls = 0;
      let weightedHits = 0;
      for (const actor of actors) {
        const hits = overlapSize(mask, actor.mask);
        if (hits === 0) continue;
        eligibleCount += actor.count;
        weightedHits += actor.count * hits;
        weightedCalls += actor.count * (1 + (nextCalls + remaining - hits) / remaining);
      }
      const expectedCalls = weightedCalls / eligibleCount;
      const averageHitRate = weightedHits / (remaining * eligibleCount);

      // A free reveal adds one known color, hence 1 + mean[T(U\c)]. The value
      // is the calls saved, normalized by the one call needed for a known color.
      // Singleton value becomes zero through this recurrence, without an override.
      const afterReveal = 1 + nextCalls / remaining;
      const difference = expectedCalls - afterReveal;
      states.push({
        mask,
        colors,
        remaining,
        expectedCalls,
        afterReveal,
        value: Math.abs(difference) < EPSILON ? 0 : difference,
        eligibleCount,
        averageHitRate,
      });
    }

    // All 127 nonempty subsets are cheap to evaluate; extrema at adjacent sizes
    // need not form a nested reveal course. Means weight each subset equally.
    const points = [];
    for (let remaining = COLORS.length; remaining >= 1; remaining -= 1) {
      const patterns = states.filter((state) => state.remaining === remaining)
        .sort((left, right) => compareMasks(left.mask, right.mask));
      let maximum = patterns[0];
      let minimum = patterns[0];
      let valueSum = 0;
      let beforeSum = 0;
      let afterSum = 0;
      for (const state of patterns) {
        if (state.value > maximum.value + EPSILON) maximum = state;
        if (state.value < minimum.value - EPSILON) minimum = state;
        valueSum += state.value;
        beforeSum += state.expectedCalls;
        afterSum += state.afterReveal;
      }
      points.push({
        remaining,
        maximum: maximum.value,
        minimum: minimum.value,
        average: valueSum / patterns.length,
        maximumColors: [...maximum.colors],
        minimumColors: [...minimum.colors],
        patterns: patterns.length,
        maximumBefore: maximum.expectedCalls,
        maximumAfter: maximum.afterReveal,
        minimumBefore: minimum.expectedCalls,
        minimumAfter: minimum.afterReveal,
        averageBefore: beforeSum / patterns.length,
        averageAfter: afterSum / patterns.length,
      });
    }

    return { ...metadata, status: "ready", points, states };
  }

  const api = Object.freeze({ analyze });
  root.NijinanaColorValue = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(globalThis);
