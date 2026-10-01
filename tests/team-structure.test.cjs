const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const source = fs.readFileSync(path.join(__dirname, '../app-export.js'), 'utf8');
const context = vm.createContext({
  CATEGORIES: ['Tecniche', 'Attitudinali', 'Sociali'], CATEGORY_MAP_NORMALIZED: new Map()
});
vm.runInContext(source.slice(source.indexOf('  function analyzeData('), source.indexOf('  function renderTeamSummary(')), context);
const { computeTeamStructure, analyzeData } = context;
function graph(n, edges) {
  const matrix = Array.from({ length: n }, () => Array(n).fill(0));
  edges.forEach(([a, b]) => { matrix[a][b] = matrix[b][a] = 1; });
  return matrix;
}
const plain = value => JSON.parse(JSON.stringify(value));

test('two triangles joined by one bridge form two communities but one component', () => {
  const matrix = graph(6, [[0,1], [1,2], [0,2], [3,4], [4,5], [3,5], [2,3]]);
  const before = JSON.stringify(matrix);
  const result = computeTeamStructure(matrix);
  assert.deepEqual(plain(result.communities), [[0,1,2], [3,4,5]]);
  assert.equal(result.components.length, 1);
  assert.equal(result.largest, 6);
  assert.equal(result.diameter, 3);
  assert.equal(result.isolated.length, 0);
  assert.equal(JSON.stringify(matrix), before);
});

test('disconnected pairs and one isolate: no misleading finite whole-team diameter', () => {
  const result = computeTeamStructure(graph(5, [[0,1], [2,3]]));
  assert.equal(result.components.length, 3);
  assert.equal(result.communities.length, 2);
  assert.equal(result.largest, 2);
  assert.equal(result.diameter, null);
  assert.deepEqual(plain(result.isolated), [4]);
});

test('empty, single-player and no-edge teams distinguish missing diameter and communities', () => {
  for (const n of [0,1,4]) {
    const result = computeTeamStructure(graph(n, []));
    assert.equal(result.communities.length, 0);
    assert.equal(result.isolated.length, n);
    assert.equal(result.diameter, null);
    assert.equal(result.largest, n ? 1 : 0);
  }
});

test('a complete team graph stays one community, path diameter uses shortest paths', () => {
  const complete = Array.from({ length: 5 }, (_, i) => Array.from({ length: 5 }, (_, j) => i === j ? 0 : 1));
  const result = computeTeamStructure(complete);
  assert.equal(result.communities.length, 1);
  assert.equal(result.diameter, 1);
  assert.equal(computeTeamStructure(graph(5, [[0,1], [1,2], [2,3], [3,4]])).diameter, 4);
});

test('category matrices never create reciprocity by mixing social and technical nominations', () => {
  const headers = ['Timestamp', 'Seleziona il tuo nome', '[Sociali] Positiva', '[Sociali] Negativa', '[Tecniche] Positiva', '[Tecniche] Negativa'];
  const analysis = analyzeData(headers, [
    ['', 'A', 'B', '', '', ''], ['', 'B', '', '', 'A', ''], ['', 'C', '', '', '', '']
  ]);
  assert.equal(computeTeamStructure(analysis.matrices.posMatrix).isolated.length, 1);
  for (const category of ['Sociali', 'Tecniche', 'Attitudinali']) {
    assert.equal(computeTeamStructure(analysis.matrices.positiveByCategory[category]).isolated.length, 3);
  }
});

test('one-way and self nominations do not create mutual graph edges', () => {
  const result = computeTeamStructure([[5,1,0], [0,0,2], [0,1,9]]);
  assert.deepEqual(plain(result.isolated), [0]);
  assert.deepEqual(plain(result.communities), [[1,2]]);
});

test('all four-player graphs retain every player exactly once and keep communities within components', () => {
  const edges = [[0,1], [0,2], [0,3], [1,2], [1,3], [2,3]];
  for (let mask = 0; mask < 64; mask += 1) {
    const result = computeTeamStructure(graph(4, edges.filter((_, i) => mask & (1 << i))));
    assert.deepEqual(plain(result.communities.flat().concat(result.isolated).sort()), [0,1,2,3]);
    for (const community of result.communities) {
      assert.ok(result.components.some(component => community.every(i => component.includes(i))));
    }
    assert.equal(result.diameter === null, result.components.length !== 1);
  }
});

test('embedded standalone app and stylesheet match maintained sources', () => {
  const app = fs.readFileSync(path.join(__dirname, '../app.js'), 'utf8');
  const css = fs.readFileSync(path.join(__dirname, '../styles.css'), 'utf8');
  const embeddedApp = app.match(/^  const INLINE_APP = (.*);$/m)[1];
  const embeddedCss = app.match(/^  const INLINE_STYLES = (.*);$/m)[1];
  assert.equal(JSON.parse(embeddedApp), source);
  assert.equal(JSON.parse(embeddedCss), css);
});

test('plain-language reading distinguishes connected subgroups, disconnected groups and absent evidence', () => {
  const describe = (n, edges) => context.describeTeamStructure(computeTeamStructure(graph(n, edges)), Array.from({length:n}, (_, i) => `Atleta ${i}`));
  assert.match(describe(1, []).title, /insufficienti/);
  assert.match(describe(3, []).title, /Non emergono/);
  const separated = describe(5, [[0,1], [2,3]]);
  assert.match(separated.title, /gruppi senza collegamenti/);
  assert.match(separated.explanation, /Atleta 4/);
  assert.match(describe(3, [[0,1]]).title, /non coinvolgono tutte/);
  const connected = describe(6, [[0,1], [0,2], [1,2], [2,3], [3,4], [3,5], [4,5]]);
  assert.match(connected.title, /Tutte le giocatrici sono collegate/);
  assert.match(connected.next, /non indica un problema/);
});
