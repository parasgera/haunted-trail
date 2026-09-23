// node score.test.cjs
const assert = require('assert');
const { match, board } = require('./site/score.js');

assert(match('  A Keyboard!! ', 'keyboard'));
assert(match('git revert', 'revert|git revert'));
assert(match('O(log n)', 'O(log n)'));
assert(!match('', ''));
assert(!match('keyboards', 'keyboard'));

const H = 3600e3, L = [{ n: 1, open: 0, close: 10 * H }, { n: 2, open: 24 * H, close: 34 * H }];
const keys = { 1: ['keyboard'], 2: ['Stack', 'October'] };
const sub = (uid, level, answers, at) => ({ uid, name: uid, level, answers: JSON.stringify(answers), at });
const b = board(L, keys, [
  sub('ann', 1, ['nope'], 1 * H),
  sub('ann', 1, ['keyboard'], 5 * H),      // latest in window wins -> correct, half-window speed bonus
  sub('ann', 2, ['stack', 'october'], 24 * H), // instant -> full speed bonus + streak 2
  sub('bob', 1, ['keyboard'], 11 * H),     // after close (edited) -> void
  sub('bob', 2, ['stack', 'june'], 34 * H), // 1/2 = cleared, zero speed bonus, streak 1
  sub('eve', 3, ['x'], 1),                 // unknown level
  { uid: 'mal', name: 'mal', level: 1, answers: '{bad json', at: 1 },
]);
assert.deepStrictEqual(b.map(e => e.id), ['ann', 'bob']);
assert.strictEqual(b[0].p, (100 + 25) + (200 + 50 + 20));
assert.deepStrictEqual([b[0].s, b[0].b], [2, 2]);
assert.strictEqual(b[1].p, 100);
assert.strictEqual(b[1].r[1], undefined);

// unrevealed level never scores even with a correct-looking answer
assert.strictEqual(board(L, { 1: ['keyboard'] }, [sub('x', 2, ['Stack', 'October'], 25 * H)]).length, 0);
console.log('score ok');
