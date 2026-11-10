import { createHash, randomUUID } from 'node:crypto';
import { lstat, mkdir, readFile, readdir, rename, rm } from 'node:fs/promises';
import path from 'node:path';

const hash = value => createHash('sha256').update(value).digest('hex');
const safeName = value => (value.normalize('NFKC').replace(/[^\p{L}\p{N}._-]+/gu, '-').replace(/^-+|-+$/g, '') || 'presentation').slice(0, 96);
const missing = error => error.code === 'ENOENT';

export function exportFingerprint(project, workspace, fileName) {
  return hash(JSON.stringify({ workspace, fileName, entry: project.source.entryName,
    manifest: project.source.manifest, pages: [...project.source.pages].sort(([a], [b]) => a.localeCompare(b)),
    assets: [...project.source.assets.values()].map(asset => [asset.path, hash(asset.bytes)]).sort(([a], [b]) => a.localeCompare(b)) }));
}

/** Pending receipts recover publication without replaying an old session snapshot. Caller holds the session lock. */
export class PptExports {
  constructor(store) { this.store = store; }
  directory(sessionId) { return path.join(this.store.sessionDirectory(sessionId), 'exports'); }
  receiptPath(sessionId, id) { return path.join(this.directory(sessionId), `${id}.json`); }
  async receipts(sessionId) {
    let names;
    try { names = await readdir(this.directory(sessionId)); }
    catch (error) { if (missing(error)) return []; throw error; }
    return Promise.all(names.filter(name => /^[a-f0-9-]{36}\.json$/.test(name))
      .map(async name => JSON.parse(await readFile(path.join(this.directory(sessionId), name), 'utf8'))));
  }
  async find(sessionId, fingerprint, callId, state) {
    const receipts = await this.receipts(sessionId);
    const exact = callId && receipts.find(receipt => receipt.callId === callId || receipt.callIds?.includes(callId));
    if (exact && exact.fingerprint !== fingerprint) throw new Error('Export retry input has changed; use a new tool call');
    const found = exact ?? receipts.find(receipt => receipt.fingerprint === fingerprint && !receipt.completed
      && !state.decks.some(deck => deck.id === receipt.deck.id));
    if (found && callId && !exact) {
      found.callIds = [...(found.callIds ?? []), callId];
      await this.writeReceipt(sessionId, found);
    }
    return found;
  }
  async assertCapacity(sessionId, state) {
    const pending = (await this.receipts(sessionId)).filter(receipt => !receipt.completed && !state.decks.some(deck => deck.id === receipt.deck.id));
    this.store.assertDeckCapacity({ ...state, decks: [...state.decks, ...pending] });
  }
  async writeReceipt(sessionId, receipt) {
    await mkdir(this.directory(sessionId), { recursive: true, mode: 0o700 });
    await this.store.writeJson(this.receiptPath(sessionId, receipt.id), receipt);
  }
  async reserve(sessionId, fingerprint, callId, deck, bytes, workspace, fileName, activity) {
    const id = randomUUID();
    const base = safeName(deck.title);
    const occupied = new Set((await this.receipts(sessionId)).map(receipt => receipt.deck.output.workspaceDirectoryPath));
    let target;
    for (let suffix = 1; suffix <= 10000; suffix++) {
      const candidate = path.join(workspace, suffix === 1 ? base : `${base}-${suffix}`);
      if (occupied.has(candidate)) continue;
      try { await lstat(candidate); }
      catch (error) { if (missing(error)) { target = candidate; break; } throw error; }
    }
    if (!target) throw new Error(`DSH PPT workspace has too many outputs named ${base}`);
    const visibleName = `${safeName(path.basename(fileName, path.extname(fileName)))}.pptx`;
    const storageKey = path.join('sessions', path.basename(this.store.sessionDirectory(sessionId)), 'outputs', `${deck.id}.pptx`);
    deck = { ...deck, output: { ...deck.output, storageKey, fileName: visibleName,
      workspaceDirectoryPath: target, workspaceFilePath: path.join(target, visibleName) } };
    const receipt = { version: 1, id, fingerprint, ...(callId ? { callId } : {}), deck, activity, completed: false };
    const internal = path.join(this.store.root, storageKey);
    await mkdir(path.dirname(internal), { recursive: true, mode: 0o700 });
    await this.store.writeFile(internal, bytes, 0o600);
    try { await this.writeReceipt(sessionId, receipt); }
    catch (error) { await rm(internal, { force: true }).catch(() => {}); throw error; }
    return receipt;
  }
  async verifyPublished(receipt, project) {
    const output = receipt.deck.output;
    const directory = await lstat(output.workspaceDirectoryPath);
    if (!directory.isDirectory() || directory.isSymbolicLink()) throw new Error('Export destination is no longer an independent directory');
    const check = async (relative, expected) => {
      const target = path.join(output.workspaceDirectoryPath, relative);
      // Refuse symlinks in any component; retries never replace user-modified output.
      let current = output.workspaceDirectoryPath;
      for (const part of relative.split('/')) {
        current = path.join(current, part);
        if ((await lstat(current)).isSymbolicLink()) throw new Error('Published export contains a symbolic link');
      }
      if (hash(await readFile(target)) !== expected) throw new Error('Published export has changed; preserve it and use a new output request');
    };
    const marker = JSON.parse(await readFile(path.join(output.workspaceDirectoryPath, '.dsh-ppt-export.json'), 'utf8'));
    if (marker.operationId !== receipt.id) throw new Error('Export destination belongs to another operation');
    await check(output.fileName, output.sha256);
    await check(project.source.entryName, hash(project.source.manifest));
    for (const [file, content] of project.source.pages) await check(file, hash(content));
    for (const asset of project.source.assets.values()) await check(asset.path, hash(asset.bytes));
  }
  async retarget(sessionId, receipt) {
    const output = receipt.deck.output, workspace = path.dirname(output.workspaceDirectoryPath);
    const base = safeName(receipt.deck.title);
    for (let suffix = 1; suffix <= 10000; suffix++) {
      const target = path.join(workspace, suffix === 1 ? base : `${base}-${suffix}`);
      try { await lstat(target); }
      catch (error) {
        if (!missing(error)) throw error;
        receipt.deck = { ...receipt.deck, output: { ...output, workspaceDirectoryPath: target, workspaceFilePath: path.join(target, output.fileName) } };
        await this.writeReceipt(sessionId, receipt);
        return;
      }
    }
    throw new Error(`DSH PPT workspace has too many outputs named ${base}`);
  }
  async publish(sessionId, receipt, project) {
    let existing;
    try { existing = await lstat(receipt.deck.output.workspaceDirectoryPath); }
    catch (error) { if (!missing(error)) throw error; }
    if (existing) {
      let marker;
      if (existing.isDirectory() && !existing.isSymbolicLink()) {
        try { marker = JSON.parse(await readFile(path.join(receipt.deck.output.workspaceDirectoryPath, '.dsh-ppt-export.json'), 'utf8')); }
        catch (error) { if (!missing(error) && !(error instanceof SyntaxError)) throw error; }
      }
      if (marker?.operationId === receipt.id) { await this.verifyPublished(receipt, project); return; }
      await this.retarget(sessionId, receipt);
    }
    const output = receipt.deck.output;
    const bytes = await readFile(path.join(this.store.root, output.storageKey));
    if (hash(bytes) !== output.sha256) throw new Error('Stored export is incomplete; cannot recover publication');
    const stage = `${output.workspaceDirectoryPath}.${receipt.id}.stage`;
    try {
      await mkdir(stage, { mode: 0o700 });
    } catch (error) {
      if (error.code !== 'EEXIST') throw error;
      const info = await lstat(stage);
      if (!info.isDirectory() || info.isSymbolicLink()) throw new Error('Export stage is invalid');
      await rm(stage, { recursive: true });
      await mkdir(stage, { mode: 0o700 });
    }
    try {
      await this.store.writeProjectFile(stage, project.source.entryName, project.source.manifest);
      for (const [file, content] of project.source.pages) await this.store.writeProjectFile(stage, file, content);
      for (const asset of project.source.assets.values()) await this.store.writeProjectFile(stage, asset.path, asset.bytes);
      await this.store.writeFile(path.join(stage, output.fileName), bytes, 0o644);
      await this.store.writeFile(path.join(stage, '.dsh-ppt-export.json'), JSON.stringify({ operationId: receipt.id }), 0o600);
      for (;;) {
        try { await mkdir(receipt.deck.output.workspaceDirectoryPath, { mode: 0o755 }); break; }
        catch (error) { if (error.code !== 'EEXIST') throw error; await this.retarget(sessionId, receipt); }
      }
      await rename(stage, receipt.deck.output.workspaceDirectoryPath);
    } finally { await rm(stage, { recursive: true, force: true }).catch(() => {}); }
  }
  async complete(sessionId, receipt) {
    // Session state is already committed; a failed acknowledgement is recovered from its deck ID.
    await this.writeReceipt(sessionId, { ...receipt, completed: true }).catch(() => {});
  }
}
