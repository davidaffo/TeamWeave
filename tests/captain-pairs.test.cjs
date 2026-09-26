const { test } = require('node:test');
const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const { join } = require('node:path');
const vm = require('node:vm');

// Exercise the same calculation shipped in the standalone application.
const source = readFileSync(join(__dirname, '..', 'app-export.js'), 'utf8');
const context = vm.createContext({});
vm.runInContext(source.slice(source.indexOf('  function analyzeCaptainPairs('), source.indexOf('  function renderCaptainPairSection(')), context);
vm.runInContext(source.slice(source.indexOf('  function buildReciprociMatrix('), source.indexOf('  function buildNonRicambiateMatrix(')), context);
vm.runInContext(source.slice(source.indexOf('  function createMatrix('), source.indexOf('  function createRectMatrix(')), context);
const { analyzeCaptainPairs, rankCaptainPairs } = context;

function fixture(size, edges) {
  const names = Array.from({ length: size }, (_, index) => `Atleta ${index}`);
  const matrix = names.map(() => names.map(() => 0));
  edges.forEach(([a, b]) => { matrix[a][b] = matrix[b][a] = 1; });
  return {
    names,
    matrices: { leadershipPos: matrix, leadershipNeg: names.map(() => names.map(() => 0)) },
    classifications: names.map(() => ({ byCategory: { Attitudinali: { inflPos: "Media" }, Sociali: { inflPos: "Media" } } })),
    summaryRows: names.map(() => ({ positiveReceived: { Attitudinali: 0 }, negativeReceived: { Attitudinali: 0 } }))
  };
}

test('shared teammates count once in coverage; the candidates themselves are excluded', () => {
  const analysis = fixture(5, [[0, 1], [0, 2], [1, 2], [0, 3], [1, 4]]);
  const pair = analyzeCaptainPairs(analysis).find((entry) => entry.first === 0 && entry.second === 1);
  assert.equal(pair.covered, 3);
  assert.equal(pair.totalPeers, 3);
  assert.equal(pair.shared.length, 1);
  assert.equal(pair.onlyFirst[0], 3);
  assert.equal(pair.onlySecond[0], 4);
  assert.equal(pair.degreeFirst, 2);
  assert.equal(pair.degreeSecond, 2);
});

test('coverage breaks equal profile scores on every graph of five athletes', () => {
  const possibleEdges = [];
  for (let a = 0; a < 5; a += 1) for (let b = a + 1; b < 5; b += 1) possibleEdges.push([a, b]);
  for (let mask = 0; mask < 1024; mask += 1) {
    const edges = possibleEdges.filter((_, index) => mask & (1 << index));
    const pairs = analyzeCaptainPairs(fixture(5, edges));
    assert.equal(pairs.length, 10);
    for (const pair of pairs) {
      const candidates = new Set([pair.first, pair.second]);
      const covered = new Set();
      let sum = 0;
      edges.forEach(([a, b]) => {
        if (candidates.has(a)) { sum += 1; if (!candidates.has(b)) covered.add(b); }
        if (candidates.has(b)) { sum += 1; if (!candidates.has(a)) covered.add(a); }
      });
      assert.equal(pair.covered, covered.size);
      assert.equal(pair.degreeFirst + pair.degreeSecond, sum - 2 * (edges.some(([a, b]) => a === pair.first && b === pair.second) ? 1 : 0));
      assert.equal(pair.covered + pair.uncovered.length, 3);
    }
    assert.equal(rankCaptainPairs(pairs)[0].covered, Math.max(...pairs.map((pair) => pair.covered)));
  }
});

test('empty teams and pairs without other teammates', () => {
  assert.equal(analyzeCaptainPairs(fixture(0, [])).length, 0);
  assert.equal(analyzeCaptainPairs(fixture(1, [])).length, 0);
  const [pair] = analyzeCaptainPairs(fixture(2, [[0, 1]]));
  assert.equal(pair.covered, 0);
  assert.equal(pair.totalPeers, 0);
  assert.equal(pair.degreeFirst, 0);
  assert.equal(pair.degreeSecond, 0);
});


test('strong attitude and social profiles take precedence over greater coverage', () => {
  const analysis = fixture(5, [[0, 2], [0, 3], [0, 4], [1, 2], [1, 3], [1, 4]]);
  for (const index of [2, 3]) {
    analysis.classifications[index].byCategory.Attitudinali.inflPos = 'Alta';
    analysis.classifications[index].byCategory.Sociali.inflPos = 'Alta';
  }
  analysis.classifications[0].byCategory.Attitudinali.inflPos = 'Bassa';
  const pairs = analyzeCaptainPairs(analysis);
  const top = rankCaptainPairs(pairs)[0];
  assert.equal(top.first, 2);
  assert.equal(top.second, 3);
  assert.ok(top.covered < Math.max(...pairs.map((pair) => pair.covered)));
});

test('technical statistics never affect the leadership ranking', () => {
  const analysis = fixture(4, [[0, 2], [1, 3]]);
  const before = JSON.stringify(rankCaptainPairs(analyzeCaptainPairs(analysis)));
  analysis.summaryRows.forEach((row) => { row.positiveReceived.Tecniche = 999; });
  analysis.matrices.reciprociPos = analysis.names.map(() => analysis.names.map(() => 1));
  assert.equal(JSON.stringify(rankCaptainPairs(analyzeCaptainPairs(analysis))), before);
});


test('the relationship between candidates does not affect their pair score', () => {
  const analysis = fixture(4, [[0, 2], [1, 3]]);
  const pair = () => analyzeCaptainPairs(analysis).find((row) => row.first === 0 && row.second === 1);
  const before = JSON.stringify(pair());
  analysis.matrices.leadershipPos[0][1] = 100;
  analysis.matrices.leadershipPos[1][0] = 100;
  analysis.matrices.leadershipNeg[0][1] = 100;
  analysis.matrices.leadershipNeg[1][0] = 100;
  assert.equal(JSON.stringify(pair()), before);
});

test('equal profiles and coverage stay tied regardless of other relationship metrics', () => {
  const base = { attitudeFloor: 3, socialFloor: 3, profileSum: 12, covered: 4 };
  const first = { ...base, reciprocalSum: 0, externalStrength: 0, attitudeBalance: 0 };
  const second = { ...base, reciprocalSum: 100, externalStrength: 100, attitudeBalance: 100 };
  assert.equal(rankCaptainPairs([first, second])[0], first);
  assert.equal(rankCaptainPairs([second, first])[0], second);
});

vm.runInContext(source.slice(source.indexOf('  function analyzeCouncil('), source.indexOf('  function renderCouncilSection(')), context);
const { analyzeCouncil } = context;

test('council coverage is a union of external teammates for any member count', () => {
  const analysis = fixture(6, [[0, 3], [1, 3], [2, 3], [0, 4], [2, 5], [0, 1], [1, 2]]);
  const result = analyzeCouncil(analysis, [0, 1, 2, 2]);
  assert.equal(result.members.length, 3);
  assert.equal(result.covered, 3);
  assert.equal(result.totalPeers, 3);
  assert.equal(result.peers.find((peer) => peer.index === 3).reachedBy.length, 3);
  assert.equal(result.profileSum, 12);
  const before = JSON.stringify(result);
  analysis.matrices.leadershipPos[0][1] = 99;
  analysis.matrices.leadershipPos[1][0] = 99;
  assert.equal(JSON.stringify(analyzeCouncil(analysis, [0, 1, 2])), before);
});

test('a council of two has the same profile and coverage as the existing pair analysis', () => {
  const analysis = fixture(5, [[0, 1], [0, 2], [1, 2], [0, 3], [1, 4]]);
  for (const pair of analyzeCaptainPairs(analysis)) {
    const council = analyzeCouncil(analysis, [pair.first, pair.second]);
    for (const key of ['covered', 'totalPeers', 'attitudeFloor', 'socialFloor', 'profileSum']) assert.equal(council[key], pair[key]);
  }
});

test('captain alone, all members selected, and missing influence are handled', () => {
  const analysis = fixture(3, [[0, 1], [1, 2]]);
  assert.equal(analyzeCouncil(analysis, [0]).covered, 1);
  const everyone = analyzeCouncil(analysis, [0, 1, 2]);
  assert.equal(everyone.totalPeers, 0);
  assert.equal(everyone.covered, 0);
  analysis.classifications[2].byCategory.Attitudinali.inflPos = 'Assente';
  assert.equal(analyzeCouncil(analysis, [0, 1, 2]).attitudeFloor, 0);
});

const { councilCandidates } = context;

test('complete councils enumerate every combination of the requested size exactly once', () => {
  const analysis = fixture(6, [[0, 3], [1, 3], [2, 4], [2, 5]]);
  for (let size = 1; size <= 6; size += 1) {
    for (const fixed of [null, 0, 3]) {
      const expected = [];
      for (let mask = 1; mask < 64; mask += 1) {
        const members = analysis.names.map((_, index) => index).filter((index) => mask & (1 << index));
        if (members.length === size && (fixed === null || members.includes(fixed))) expected.push(members.join(','));
      }
      const candidates = [...councilCandidates(analysis, size, fixed)];
      const actual = candidates.map((candidate) => [...candidate.members].sort((a, b) => a - b).join(','));
      assert.deepEqual(actual.sort(), expected.sort());
      assert.equal(new Set(actual).size, actual.length);
    }
  }
  assert.equal([...councilCandidates(analysis, 0)].length, 0);
  assert.equal([...councilCandidates(analysis, 7)].length, 0);
});

test('recommendations optimize the complete requested council and honor a fixed captain', () => {
  const analysis = fixture(6, [[0, 3], [2, 3], [2, 5], [4, 1]]);
  for (const index of [0, 2, 4]) {
    analysis.classifications[index].byCategory.Attitudinali.inflPos = 'Alta';
    analysis.classifications[index].byCategory.Sociali.inflPos = 'Alta';
  }
  const best = rankCaptainPairs([...councilCandidates(analysis, 3)])[0];
  assert.equal([...best.members].sort().join(','), '0,2,4');
  assert.equal(best.covered, 3);
  const fixed = rankCaptainPairs([...councilCandidates(analysis, 3, 1)])[0];
  assert.equal(fixed.members.length, 3);
  assert.ok(fixed.members.includes(1));
});
