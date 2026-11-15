const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const { resolve } = require('node:path');
const { test } = require('node:test');

const workflow = readFileSync(resolve(__dirname, '..', 'workflows', 'review-approval-policy.yml'), 'utf8');
const signal = readFileSync(resolve(__dirname, '..', 'workflows', 'review-changed.yml'), 'utf8');
const block = workflow.match(/          script: \|\r?\n((?:            [^\r\n]*(?:\r?\n|$)|\r?\n)+)/);
assert.ok(block, 'Expected an embedded approval policy script.');
const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
const evaluate = new AsyncFunction('github', 'context', 'core', block[1].replace(/^            /gm, ''));

function pull(number = 1, labels = [], sha = `sha-${number}`) {
  return {
    number, draft: false, head: { sha },
    user: { login: 'author' }, labels: labels.map(name => ({ name }))
  };
}

function review(id, login, state = 'APPROVED', type = 'User') {
  return { id, state, user: { login, type } };
}

function harness(pulls, reviews = {}, options = {}) {
  const statuses = [];
  const logs = [];
  const permissionCalls = [];
  const list = async () => {};
  const listReviews = async () => {};
  const github = {
    rest: {
      pulls: { list, listReviews },
      repos: {
        createCommitStatus: async status => {
          statuses.push(status);
        },
        getCollaboratorPermissionLevel: async ({ username }) => {
          permissionCalls.push(username);
          if (options.permissionError) {
            throw options.permissionError;
          }
          return { data: { permission: options.permissions?.[username] ?? 'write' } };
        }
      }
    },
    paginate: async (method, params) => {
      assert.equal(params.per_page, 100);
      if (method === list) {
        assert.equal(params.base, 'master');
        assert.equal(params.state, 'open');
        return pulls;
      }
      assert.equal(method, listReviews);
      if (options.reviewError) {
        throw options.reviewError;
      }
      return [...(reviews[params.pull_number] ?? [])];
    }
  };
  const context = {
    repo: { owner: 'mini-software', repo: 'MiniExcel' },
    serverUrl: 'https://github.com', runId: 123,
    payload: { workflow_run: { pull_requests: [], head_repository: { full_name: 'untrusted/fork' } } }
  };
  const run = () => evaluate(github, context, { info: message => logs.push(message) });
  return { run, statuses, logs, permissionCalls };
}

test('ordinary PR requires exactly one eligible approval', async () => {
  for (const [reviews, expected] of [[[], 'pending'], [[review(1, 'reviewer')], 'success']]) {
    const h = harness([pull()], { 1: reviews });
    await h.run();
    assert.equal(h.statuses[0].state, 'pending');
    assert.equal(h.statuses.at(-1).state, expected);
    assert.equal(h.statuses.at(-1).context, 'Review approval policy');
  }
});

test('major-change PR remains blocked at two and passes at three approvals', async () => {
  for (const count of [0, 1, 2, 3]) {
    const h = harness([pull(1, ['MAJOR-CHANGE'])], {
      1: Array.from({ length: count }, (_, index) => review(index + 1, `reviewer-${index}`))
    });
    await h.run();
    assert.equal(h.statuses.at(-1).state, count === 3 ? 'success' : 'pending');
    if (count < 3) {
      assert.match(h.statuses.at(-1).description, new RegExp(`${count}/3`));
    }
  }
});

test('invited users, comments, author approvals, bots and Read users do not count', async () => {
  const h = harness([pull(1, ['major-change'])], {
    1: [
      review(1, 'AUTHOR'), review(2, 'bot', 'APPROVED', 'Bot'),
      review(3, 'read-user'), review(4, 'commenter', 'COMMENTED'),
      review(5, 'valid')
    ]
  }, { permissions: { 'read-user': 'read' } });
  await h.run();
  assert.match(h.statuses.at(-1).description, /1\/3/);
  assert.deepEqual(h.permissionCalls, ['read-user', 'valid']);
});

test('counts Write, Maintain and Admin approvals', async () => {
  const h = harness([pull(1, ['major-change'])], {
    1: [review(1, 'writer'), review(2, 'maintainer'), review(3, 'admin')]
  }, { permissions: { writer: 'write', maintainer: 'maintain', admin: 'admin' } });
  await h.run();
  assert.equal(h.statuses.at(-1).state, 'success');
});

test('later dismissed or changes-requested reviews supersede earlier approvals', async () => {
  for (const state of ['DISMISSED', 'CHANGES_REQUESTED']) {
    const h = harness([pull()], {
      1: [review(2, 'REVIEWER', state), review(1, 'reviewer')]
    });
    await h.run();
    assert.match(h.statuses.at(-1).description, /0\/1/);
  }
});

test('comments do not cancel approvals and duplicate approvals count only once', async () => {
  const h = harness([pull(1, ['major-change'])], {
    1: [review(1, 'reviewer'), review(2, 'REVIEWER'), review(3, 'Reviewer', 'COMMENTED')]
  });
  await h.run();
  assert.match(h.statuses.at(-1).description, /1\/3/);
});

test('a new approval after a change request is counted', async () => {
  const h = harness([pull()], {
    1: [review(1, 'reviewer', 'CHANGES_REQUESTED'), review(2, 'reviewer')]
  });
  await h.run();
  assert.equal(h.statuses.at(-1).state, 'success');
});

test('a label change raises the requirement without trusting old success', async () => {
  const current = pull();
  const h = harness([current], { 1: [review(1, 'reviewer')] });
  await h.run();
  assert.equal(h.statuses.at(-1).state, 'success');
  current.labels.push({ name: 'major-change' });
  await h.run();
  assert.equal(h.statuses.at(-2).state, 'pending');
  assert.match(h.statuses.at(-1).description, /1\/3/);
});

test('PRs sharing a SHA cannot overwrite a stricter PR status with ordinary success', async () => {
  const h = harness([pull(1, [], 'shared'), pull(2, ['major-change'], 'shared')], {
    1: [review(1, 'one')],
    2: [review(2, 'one'), review(3, 'two')]
  });
  await h.run();
  assert.equal(h.statuses.length, 2);
  assert.equal(h.statuses.at(-1).sha, 'shared');
  assert.match(h.statuses.at(-1).description, /#2: 2\/3/);
  assert.deepEqual(h.permissionCalls, ['one', 'two']);
});

test('all PRs sharing a SHA must satisfy their own approval requirements', async () => {
  const h = harness([pull(1, [], 'shared'), pull(2, ['major-change'], 'shared')], {
    1: [review(1, 'one')],
    2: [review(2, 'one'), review(3, 'two'), review(4, 'three')]
  });
  await h.run();
  assert.equal(h.statuses.at(-1).state, 'success');
});

test('review and permission API failures leave checks pending and propagate', async () => {
  for (const option of ['reviewError', 'permissionError']) {
    const error = new Error('GitHub API access failed');
    const h = harness([pull()], { 1: [review(1, 'reviewer')] }, { [option]: error });
    await assert.rejects(h.run(), error);
    assert.deepEqual(h.statuses.map(status => status.state), ['pending']);
  }
});

test('draft PRs are skipped and no open PRs are reported explicitly', async () => {
  const draft = pull();
  draft.draft = true;
  for (const pulls of [[], [draft]]) {
    const h = harness(pulls);
    await h.run();
    assert.deepEqual(h.statuses, []);
    assert.match(h.logs[0], /No open/);
  }
});

test('commit status descriptions obey the GitHub length limit', async () => {
  const h = harness(Array.from({ length: 20 }, (_, index) => pull(index + 1, ['major-change'], 'shared')));
  await h.run();
  assert.equal(h.statuses.at(-1).description.length, 140);
});

test('trusted workflow uses no PR code, artifact content, team credential or write-capable review trigger', () => {
  assert.match(workflow, /pull_request_target:/);
  assert.match(workflow, /workflow_run:/);
  assert.match(workflow, /workflows: \[PR review changed\]/);
  assert.match(workflow, /statuses: write/);
  assert.match(workflow, /pull-requests: read/);
  assert.doesNotMatch(workflow, /actions\/checkout|\$\{\{[^}]*head\.(?:sha|ref)|download-artifact|secrets\.|pull_request_review:/);
  assert.match(signal, /pull_request_review:/);
  assert.match(signal, /types: \[submitted, dismissed\]/);
  assert.match(signal, /permissions: \{\}/);
  assert.doesNotMatch(signal, /actions\/checkout|\$\{\{|secrets\.|write/);
});
