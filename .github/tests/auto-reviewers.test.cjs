const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const { resolve } = require('node:path');
const { test } = require('node:test');

const workflow = readFileSync(resolve(__dirname, '..', 'workflows', 'auto-reviewers.yml'), 'utf8');
const blocks = [...workflow.matchAll(/          script: \|\r?\n((?:            [^\r\n]*(?:\r?\n|$)|\r?\n)+)/g)];
assert.equal(blocks.length, 2, 'Expected the team-reading and reviewer-requesting scripts.');
const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
const scripts = blocks.map(match => new AsyncFunction(
  'github', 'context', 'core', 'process',
  match[1].replace(/^            /gm, '')
));
const context = {
  repo: { owner: 'mini-software', repo: 'MiniExcel' },
  payload: { pull_request: { number: 1040 } }
};

function harness(options = {}) {
  const logs = [];
  const core = {
    info: message => logs.push(message),
    notice: message => logs.push(message)
  };
  const requested = (options.requested ?? []).map(login => ({ login }));
  const reviews = (options.reviewed ?? []).map(login => ({ user: { login } }));
  const batches = [];
  const pull = {
    state: 'open',
    draft: false,
    base: { ref: 'master' },
    user: { login: 'author' },
    ...options.pull
  };
  const listReviews = async () => {};
  const github = {
    rest: {
      pulls: {
        get: async params => {
          assert.ok(Number.isSafeInteger(params.pull_number));
          return { data: pull };
        },
        listRequestedReviewers: async () => ({ data: { users: requested } }),
        listReviews,
        requestReviewers: async ({ reviewers }) => {
          if (options.error) {
            throw options.error;
          }
          batches.push(reviewers);
          if (options.omit) {
            return;
          }
          if (options.reviewImmediately) {
            reviews.push(...reviewers.map(login => ({ user: { login } })));
          } else {
            requested.push(...reviewers.map(login => ({ login })));
          }
        }
      }
    },
    paginate: async (method, params) => {
      assert.equal(method, listReviews);
      assert.equal(params.per_page, 100);
      return [...reviews];
    }
  };
  const run = (members, customContext = context, input = '') => scripts[1](
    github, customContext, core,
    { env: { TEAM_REVIEWERS: JSON.stringify(members), INPUT_PR_NUMBER: input } }
  );
  return { run, logs, batches };
}

test('team reader requires an explicit organization credential', async () => {
  await assert.rejects(scripts[0]({}, context, {}, { env: {} }), /CORE_REVIEWERS_TOKEN/);
});

test('team reader dynamically paginates active members and excludes bots', async () => {
  const listMembersInOrg = async () => {};
  const github = {
    rest: { teams: { listMembersInOrg } },
    paginate: async (method, params) => {
      assert.equal(method, listMembersInOrg);
      assert.deepEqual(params, {
        org: 'mini-software', team_slug: 'core-reviewers', per_page: 100
      });
      return [{ login: 'new-member', type: 'User' }, { login: 'bot', type: 'Bot' }];
    }
  };
  const result = await scripts[0](github, context, { info() {} }, {
    env: { CORE_REVIEWERS_TOKEN: 'test-only' }
  });
  assert.deepEqual(JSON.parse(result), ['new-member']);
  github.paginate = async () => [];
  await assert.rejects(scripts[0](github, context, {}, {
    env: { CORE_REVIEWERS_TOKEN: 'test-only' }
  }), /no active human members/);
});

test('excludes the author, requested users, reviewed users and duplicate names case-insensitively', async () => {
  const h = harness({ requested: ['requested'], reviewed: ['reviewed'] });
  await h.run(['AUTHOR', 'REQUESTED', 'REVIEWED', 'new-member', 'NEW-MEMBER']);
  assert.deepEqual(h.batches, [['new-member']]);
});

test('requests all twenty eligible members in batches of at most fifteen', async () => {
  const members = Array.from({ length: 20 }, (_, index) => `member-${index}`);
  const h = harness();
  await h.run(members);
  assert.deepEqual(h.batches.map(batch => batch.length), [15, 5]);
  assert.deepEqual(h.batches.flat(), members);
});

test('repeated events do not send duplicate review requests', async () => {
  const h = harness();
  await h.run(['first', 'second']);
  await h.run(['first', 'second', 'new-member']);
  assert.deepEqual(h.batches, [['first', 'second'], ['new-member']]);
});

test('reports when no additional reviewer is eligible', async () => {
  const h = harness();
  await h.run(['author']);
  assert.deepEqual(h.batches, []);
  assert.match(h.logs[0], /already been requested or reviewed/);
});

test('skips closed, draft and non-master PRs explicitly', async () => {
  for (const pull of [{ state: 'closed' }, { draft: true }, { base: { ref: 'other' } }]) {
    const h = harness({ pull });
    await h.run(['reviewer']);
    assert.deepEqual(h.batches, []);
    assert.match(h.logs[0], /Skipping/);
  }
});

test('manual dispatch accepts a PR number and rejects invalid inputs', async () => {
  const manualContext = { ...context, payload: {} };
  const h = harness();
  await h.run(['reviewer'], manualContext, '1040');
  assert.deepEqual(h.batches, [['reviewer']]);
  for (const input of ['', '-1', '0', '1.5', 'text', '9007199254740992']) {
    await assert.rejects(h.run(['reviewer'], manualContext, input), /positive integer/);
  }
});

test('API permission failures propagate instead of reporting success', async () => {
  const error = new Error('Resource not accessible by integration');
  const h = harness({ error });
  await assert.rejects(h.run(['reviewer']), error);
  assert.deepEqual(h.logs, []);
});

test('silently omitted API requests fail verification', async () => {
  const h = harness({ omit: true });
  await assert.rejects(h.run(['reviewer']), /did not confirm review requests for: reviewer/);
  assert.deepEqual(h.logs, []);
});

test('review completed during the request is counted as confirmed', async () => {
  const h = harness({ reviewImmediately: true });
  await h.run(['reviewer']);
  assert.match(h.logs[0], /Requested 1/);
});

test('malformed team output fails instead of treating it as an empty list', async () => {
  await assert.rejects(scripts[1](
    {
      rest: { pulls: {
        get: async () => ({ data: {
          state: 'open', draft: false, base: { ref: 'master' }, user: { login: 'author' }
        } }),
        listRequestedReviewers: async () => ({ data: { users: [] } }),
        listReviews: async () => {}
      } },
      paginate: async () => []
    },
    context, {}, { env: { TEAM_REVIEWERS: 'not-json' } }
  ), SyntaxError);
});

test('fork-capable workflow never checks out or executes PR code', () => {
  assert.match(workflow, /pull_request_target:/);
  assert.match(workflow, /pull-requests: write/);
  assert.doesNotMatch(workflow, /actions\/checkout|head\.sha|head\.ref|pull_request\.title|pull_request\.body/);
  assert.match(workflow, /github-token: \$\{\{ secrets\.CORE_REVIEWERS_TOKEN \}\}/);
});
