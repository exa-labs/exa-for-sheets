const fs = require('fs');
const path = require('path');
const vm = require('vm');

function createSidebarHarness() {
  const html = fs.readFileSync(path.join(__dirname, '..', 'Sidebar.html'), 'utf8');
  const section = html.split('// --- One Agent run per source row ---')[1]
    .split('// --- Agent Fill Logic ---')[0];
  const pendingCalls = [];
  const elements = {};
  const confirm = jest.fn(() => { throw new Error('Unexpected confirmation dialog'); });
  const runner = {
    withSuccessHandler(onSuccess) {
      return { withFailureHandler: onFailure => ({
        ensureAuthorized: () => pendingCalls.push({ method: 'authorize', onSuccess, onFailure }),
        prepareAgentForEachJobs: config => pendingCalls.push({ method: 'prepare', config, onSuccess, onFailure }),
        startAgentForEachRun: (job) => pendingCalls.push({ method: 'start', job, onSuccess, onFailure }),
        getAgentRunStatus: () => pendingCalls.push({ method: 'poll', onSuccess, onFailure }),
        writeAgentForEachRunToSheet: () => pendingCalls.push({ method: 'write', onSuccess, onFailure }),
        cancelAgentRun: () => pendingCalls.push({ method: 'cancel', onSuccess, onFailure })
      }) };
    }
  };
  const context = vm.createContext({
    google: { script: { run: runner } },
    document: { getElementById: id => (elements[id] ||= {
      value: '', textContent: '', style: {}, classList: { toggle() {} }
    }) },
    confirm,
    setTimeout,
    clearTimeout
  });
  const declarations = `
    let agentForEachQueue = [];
    let agentForEachActive = {};
    let agentForEachCompleted = 0;
    let agentForEachTotal = 0;
    let agentForEachRows = 0;
    let agentForEachFailures = [];
    let agentForEachCancelled = false;
    let agentForEachConfig = {};
    let agentForEachCost = 0;
    let agentForEachGeneration = 1;
    let agentForEachRpcQueue = [];
    let agentForEachRpcInFlight = 0;
    const AGENT_FOR_EACH_CONCURRENCY = 50;
    const AGENT_FOR_EACH_RPC_CONCURRENCY = 8;
  `;
  vm.runInContext(declarations + section, context);
  return { context, pendingCalls, elements, confirm };
}

test('clicking Run starts prepared per-row Agent calls without a second confirmation', () => {
  const { context, pendingCalls, elements, confirm } = createSidebarHarness();
  elements.agentForEachPrompt = { value: 'Find relevant people' };
  elements.agentForEachDestination = { value: '22' };
  elements.agentForEachEffort = { value: 'auto' };

  context.runAgentForEach();
  const authorization = pendingCalls.shift();
  expect(authorization.method).toBe('authorize');
  authorization.onSuccess();
  const preparation = pendingCalls.shift();
  expect(preparation.method).toBe('prepare');
  expect(preparation.config).not.toHaveProperty('scope');
  preparation.onSuccess({ success: true, jobs: [
    { sourceRow: 2, sourceValue: 'First' },
    { sourceRow: 3, sourceValue: 'Second' }
  ] });

  expect(confirm).not.toHaveBeenCalled();
  expect(elements.agentForEachScope).toBeUndefined();
  expect(pendingCalls.map(call => call.method)).toEqual(['start', 'start']);
  expect(elements.agentForEachStatus.textContent).toBe('Starting 2 Agent run(s)...');
});

test('for-each mode permits 50 active runs but only eight concurrent bridge calls', () => {
  const { context, pendingCalls } = createSidebarHarness();
  vm.runInContext(`agentForEachQueue = Array.from({length: 60}, (_, index) => ({
    sourceRow: index + 2, sourceValue: 'Row ' + index
  })); agentForEachTotal = 60; startNextAgentForEachJobs();`, context);

  expect(vm.runInContext('Object.keys(agentForEachActive).length', context)).toBe(50);
  expect(vm.runInContext('agentForEachQueue.length', context)).toBe(10);
  expect(vm.runInContext('agentForEachRpcInFlight', context)).toBe(8);
  expect(pendingCalls).toHaveLength(8);

  pendingCalls.shift().onSuccess({ success: true, run: { id: 'run-1' }, writeConfig: {} });
  expect(vm.runInContext('Object.keys(agentForEachActive).length', context)).toBe(50);
  expect(vm.runInContext('agentForEachRpcInFlight', context)).toBe(8);
  expect(pendingCalls).toHaveLength(8);
});

test('cancelling drops queued starts and cancels runs created by in-flight starts', () => {
  const { context, pendingCalls } = createSidebarHarness();
  vm.runInContext(`agentForEachQueue = Array.from({length: 50}, (_, index) => ({
    sourceRow: index + 2, sourceValue: 'Row ' + index
  })); agentForEachTotal = 50; startNextAgentForEachJobs(); cancelAgentForEach();`, context);

  expect(vm.runInContext('agentForEachRpcQueue.length', context)).toBe(0);
  pendingCalls.shift().onSuccess({ success: true, run: { id: 'run-after-cancel' } });
  expect(pendingCalls).toHaveLength(8);
  expect(pendingCalls.filter(call => call.method === 'cancel')).toHaveLength(1);
  expect(vm.runInContext('agentForEachRpcInFlight', context)).toBe(8);
});

test('task picker shows one matching action and hides irrelevant fields', () => {
  const html = fs.readFileSync(path.join(__dirname, '..', 'Sidebar.html'), 'utf8');
  const setMode = html.split('function setAgentMode(mode) {')[1]
    .split('function getAgentConfigFromForm()')[0];
  const elements = {};
  const context = vm.createContext({
    document: { getElementById: id => (elements[id] ||= { hidden: false, value: '', textContent: '' }) },
    loadAgentForEachDestinations: jest.fn()
  });
  vm.runInContext(`function setAgentMode(mode) {${setMode}`, context);

  context.setAgentMode('find-more');
  expect(elements.agentRunButton.textContent).toBe('Find more');
  expect(elements.agentCreatePanel.hidden).toBe(false);
  expect(elements.agentColumnsGroup.hidden).toBe(true);
  expect(elements.agentStartCellGroup.hidden).toBe(true);
  expect(elements.agentFindMoreHelp.hidden).toBe(false);

  context.setAgentMode('for-each');
  expect(elements.agentCreatePanel.hidden).toBe(true);
  expect(elements.agentForEachPanel.hidden).toBe(false);
  expect(context.loadAgentForEachDestinations).toHaveBeenCalledTimes(1);
});
