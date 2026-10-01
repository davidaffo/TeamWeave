const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(require('node:path').join(__dirname, '../app-export.js'), 'utf8');
const context = vm.createContext({
  PALETTE: { network: ['#fff', '#aaa', '#000'] },
  sumRows: matrix => matrix.map(row => row.reduce((a, b) => a + b, 0)),
  minMaxArray: values => ({ min: Math.min(0, ...values), max: Math.max(0, ...values) }),
  layoutGraph: nodes => nodes.forEach((node, i) => { node.x = i * 100; node.y = 100; }),
  interpolateBi: () => '#aaa', interpolateTri: () => '#aaa'
});
vm.runInContext(source.slice(source.indexOf('  function computeRelationMetrics('), source.indexOf('  function updateSummary(')), context);
vm.runInContext(source.slice(source.indexOf('  function createNetworkState('), source.indexOf('  function attachNetworkToggle(')), context);

test('mutual density and reciprocity use different denominators; weights and self ties do not inflate them', () => {
  const metrics = context.computeRelationMetrics([[8, 4, 0], [1, 0, 9], [0, 0, 0]]);
  assert.equal(metrics.mutualPairs, 1);
  assert.equal(metrics.possiblePairs, 3);
  assert.equal(metrics.directedTies, 3);
  assert.equal(metrics.mutualDensity, 1 / 3);
  assert.equal(metrics.reciprocity, 2 / 3);
});

test('undefined denominators are not displayed as zero reciprocity', () => {
  for (const matrix of [[], [[0]]]) {
    assert.equal(context.computeRelationMetrics(matrix).mutualDensity, null);
    assert.equal(context.computeRelationMetrics(matrix).reciprocity, null);
  }
  const empty = context.computeRelationMetrics([[0, 0], [0, 0]]);
  assert.equal(empty.mutualDensity, 0);
  assert.equal(empty.reciprocity, null);
});

test('fully reciprocated isolated pairs do not imply a dense team network', () => {
  const metrics = context.computeRelationMetrics([[0,1,0,0], [1,0,0,0], [0,0,0,1], [0,0,1,0]]);
  assert.equal(metrics.reciprocity, 1);
  assert.equal(metrics.mutualDensity, 1 / 3);
});

test('focus dims unrelated nodes, strength filter removes edges and canvas style resets', () => {
  const matrix = [[0, 1, 0], [1, 0, 1], [0, 1, 0]];
  const state = context.createNetworkState(['A', 'B', 'C'], matrix, [[0, 2, 0], [2, 0, 5], [0, 5, 0]], 900, 520);
  state.selectedIndex = 0;
  const labels = [], strokes = [];
  const ctx = {
    clearRect() {}, beginPath() {}, moveTo() {}, lineTo() {}, arc() {}, fill() {},
    stroke() { strokes.push(this.globalAlpha); },
    fillText(name) { labels.push([name, this.globalAlpha, this.filter]); }
  };
  context.drawNetwork(ctx, state);
  assert.deepEqual(labels, [['A', 1, 'none'], ['B', 1, 'none'], ['C', .18, 'blur(2px)']]);
  assert.equal(ctx.globalAlpha, 1);
  assert.equal(ctx.filter, 'none');
  state.minStrength = 3;
  strokes.length = 0;
  context.drawNetwork(ctx, state);
  assert.equal(strokes.length, 4); // One edge and three node outlines.
});

test('pointer selection works with actual node weights, supports clearing and does not select on drag', () => {
  const state = context.createNetworkState(['A', 'B'], [[0, 1], [1, 0]], null, 900, 520);
  const handlers = {};
  const ctx = new Proxy({}, { get: (target, key) => target[key] ?? (() => {}), set: (target, key, value) => { target[key] = value; return true; } });
  const canvas = { _networkState: state, width: 900, height: 520, getContext: () => ctx,
    getBoundingClientRect: () => ({ left: 0, top: 0, width: 900, height: 520 }),
    setPointerCapture() {}, addEventListener: (name, handler) => { handlers[name] = handler; } };
  context.attachNetworkHandlers(canvas);
  const event = (x, y) => ({ clientX: x, clientY: y, pointerId: 1 });
  handlers.pointerdown(event(0, 100)); handlers.pointerup(event(0, 100));
  assert.equal(state.selectedIndex, 0);
  handlers.pointerdown(event(500, 400)); handlers.pointerup(event(500, 400));
  assert.equal(state.selectedIndex, null);
  handlers.pointerdown(event(0, 100)); handlers.pointermove(event(50, 100)); handlers.pointerup(event(50, 100));
  assert.equal(state.selectedIndex, null);
  assert.equal(state.nodes[0].x, 50);
});
