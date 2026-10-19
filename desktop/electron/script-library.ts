// Copyright (c) 2026 AbeLin · MIT License
import { createHash, randomUUID } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, readdirSync, renameSync, statSync, unlinkSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import type { CodeResult, HistoryItem, ScriptInfo, ScriptLibrarySnapshot, ScriptRecord, ScriptSave, ScriptStatus } from '../shared/types';

const statuses: ScriptStatus[] = ['never', 'running', 'success', 'error', 'stopped', 'timeout'];
const maxBackupBytes = 64 * 1024 * 1024;

function codeText(value: unknown): string {
  if (typeof value !== 'string' || !value.trim() || value.length > 200000) throw new Error('脚本不能为空，最多 200,000 字符');
  return value.replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n');
}
function scriptId(code: string) { return createHash('sha256').update(code).digest('hex'); }
function field(value: unknown, max: number, fallback = ''): string {
  if (value === undefined) return fallback;
  if (typeof value !== 'string' || value.length > max) throw new Error(`脚本信息格式无效，字段最多 ${max} 字符`);
  return value.trim();
}
function date(value: unknown, fallback: string): string {
  return typeof value === 'string' && Number.isFinite(Date.parse(value)) ? new Date(value).toISOString() : fallback;
}
function validate(value: unknown): ScriptRecord {
  if (!value || typeof value !== 'object') throw new Error('脚本记录格式无效');
  const raw = value as Record<string, unknown>;
  const code = codeText(raw.code);
  const id = scriptId(code);
  if (raw.id !== id) throw new Error('脚本内容与记录编号不一致');
  const now = new Date().toISOString();
  return {
    id, code, name: field(raw.name, 120, '未命名脚本') || '未命名脚本', category: field(raw.category, 40, '未分类') || '未分类',
    notes: field(raw.notes, 2000), favorite: raw.favorite === true, archived: raw.archived === true,
    createdAt: date(raw.createdAt, now), updatedAt: date(raw.updatedAt, now), lastRunAt: typeof raw.lastRunAt === 'string' ? date(raw.lastRunAt, now) : null,
    runCount: typeof raw.runCount === 'number' && Number.isSafeInteger(raw.runCount) && raw.runCount >= 0 ? raw.runCount : 0,
    status: statuses.includes(raw.status as ScriptStatus) ? raw.status as ScriptStatus : 'never', lines: code.split('\n').length,
  };
}

export class ScriptLibrary {
  constructor(readonly path: string) {}

  private file(id: string) {
    if (!/^[a-f0-9]{64}$/.test(id)) throw new Error('脚本编号无效');
    return join(this.path, `${id}.json`);
  }

  private atomicWrite(path: string, content: string) {
    mkdirSync(this.path, { recursive: true });
    const temporary = `${path}.${randomUUID()}.tmp`;
    try { writeFileSync(temporary, content, { encoding: 'utf8', flag: 'wx' }); renameSync(temporary, path); }
    finally { if (existsSync(temporary)) unlinkSync(temporary); }
  }

  private write(record: ScriptRecord) {
    const checked = validate(record);
    this.atomicWrite(this.file(checked.id), JSON.stringify(checked, null, 2));
    return checked;
  }

  get(id: string): ScriptRecord {
    const path = this.file(id);
    if (!existsSync(path)) throw new Error('脚本已不存在，请刷新脚本库');
    if (statSync(path).size > 900000) throw new Error('脚本记录过大');
    const record = validate(JSON.parse(readFileSync(path, 'utf8')));
    if (record.id !== id) throw new Error('脚本记录编号不一致');
    return record;
  }

  list(): ScriptLibrarySnapshot {
    mkdirSync(this.path, { recursive: true });
    const items: ScriptInfo[] = [];
    const warnings: string[] = [];
    for (const filename of readdirSync(this.path)) {
      if (!/^[a-f0-9]{64}\.json$/.test(filename)) continue;
      try {
        const { code: _code, ...metadata } = this.get(filename.slice(0, -5));
        items.push(metadata);
      } catch { warnings.push(`有一份记录无法读取，原文件已保留：${filename}`); }
    }
    items.sort((a, b) => (b.lastRunAt || b.updatedAt).localeCompare(a.lastRunAt || a.updatedAt));
    return { items, warnings, path: this.path };
  }

  save(request: ScriptSave): ScriptRecord {
    const code = codeText(request?.code);
    const id = scriptId(code);
    const existing = existsSync(this.file(id)) ? this.get(id) : null;
    if (existing && request.rememberOnly) return existing;
    const now = new Date().toISOString();
    const base: ScriptRecord = existing || {
      id, code, createdAt: now, lastRunAt: null, runCount: 0, status: 'never', archived: false,
      name: '未命名脚本', category: '未分类', notes: '', favorite: false, lines: code.split('\n').length,
      updatedAt: now,
    };
    return this.write({
      ...base, updatedAt: now, archived: false,
      name: field(request.name, 120, existing?.name || '未命名脚本') || '未命名脚本',
      category: field(request.category, 40, existing?.category || '未分类') || '未分类',
      notes: field(request.notes, 2000, existing?.notes || ''),
      favorite: typeof request.favorite === 'boolean' ? request.favorite : existing?.favorite || false,
    });
  }

  update(id: string, changes: Partial<Pick<ScriptInfo, 'name' | 'category' | 'notes' | 'favorite' | 'archived'>>) {
    if (!changes || typeof changes !== 'object') throw new Error('脚本信息格式无效');
    const record = this.get(id);
    if (changes.name !== undefined) record.name = field(changes.name, 120) || '未命名脚本';
    if (changes.category !== undefined) record.category = field(changes.category, 40) || '未分类';
    if (changes.notes !== undefined) record.notes = field(changes.notes, 2000);
    for (const key of ['favorite', 'archived'] as const) {
      if (changes[key] !== undefined) {
        if (typeof changes[key] !== 'boolean') throw new Error('脚本标记格式无效');
        record[key] = changes[key];
      }
    }
    record.updatedAt = new Date().toISOString();
    return this.write(record);
  }

  recordRun(code: string, name = '未命名脚本', at = new Date().toISOString()) {
    const record = this.save({ code, name: name.replace(/\.py$/i, ''), rememberOnly: true });
    record.runCount += 1;
    record.lastRunAt = at;
    record.updatedAt = at;
    record.status = 'running';
    record.archived = false;
    return this.write(record);
  }

  finishRun(id: string, status: CodeResult['status']) {
    const record = this.get(id);
    record.status = status;
    record.updatedAt = new Date().toISOString();
    return this.write(record);
  }

  exportBackup(): string {
    const snapshot = this.list();
    if (snapshot.warnings.length) throw new Error('存在无法读取的脚本记录，请先修复或从之前的备份恢复，再导出完整备份');
    const backup = JSON.stringify({ format: 'ExcelRouterScripts', version: 1, exportedAt: new Date().toISOString(), scripts: snapshot.items.map((item) => this.get(item.id)) }, null, 2);
    if (Buffer.byteLength(backup, 'utf8') > maxBackupBytes) throw new Error('脚本库超过 64 MB，请将部分脚本单独另存为 .py');
    return backup;
  }

  importBackup(content: string) {
    if (Buffer.byteLength(content, 'utf8') > maxBackupBytes) throw new Error('备份最多 64 MB');
    const backup = JSON.parse(content.replace(/^\uFEFF/, ''));
    if (backup?.format !== 'ExcelRouterScripts' || backup.version !== 1 || !Array.isArray(backup.scripts) || backup.scripts.length > 10000) throw new Error('请选择 EasyExcel 或旧版 ExcelRouter 导出的脚本库备份');
    // Validate the entire document before any writes; unknown paths are never used.
    const incoming = backup.scripts.map((item: unknown) => validate(item)) as ScriptRecord[];
    const mergedRecords = new Map<string, ScriptRecord>();
    const damagedFiles = new Set<string>();
    let added = 0, merged = 0;
    for (const record of incoming) {
      let old = mergedRecords.get(record.id) || null;
      if (!old && existsSync(this.file(record.id))) {
        try { old = this.get(record.id); } catch { damagedFiles.add(this.file(record.id)); }
      }
      if (old) {
        merged += 1;
        const newerRun = (record.lastRunAt || '') > (old.lastRunAt || '');
        mergedRecords.set(record.id, { ...old, favorite: old.favorite || record.favorite,
          runCount: Math.max(old.runCount, record.runCount), lastRunAt: newerRun ? record.lastRunAt : old.lastRunAt,
          status: newerRun ? record.status : old.status });
      } else { added += 1; mergedRecords.set(record.id, record); }
    }
    // Keep a pre-import snapshot so interrupted imports can also be recovered.
    const snapshot = this.list();
    this.atomicWrite(join(this.path, 'before-import.backup'), JSON.stringify({ format: 'ExcelRouterScripts', version: 1,
      exportedAt: new Date().toISOString(), scripts: snapshot.items.map((item) => this.get(item.id)), warnings: snapshot.warnings }));
    for (const path of damagedFiles) renameSync(path, `${path}.${randomUUID()}.damaged`);
    for (const record of mergedRecords.values()) this.write(record);
    return { added, merged };
  }

  migrateHistory(history: HistoryItem[]) {
    const marker = join(this.path, 'legacy-import.state');
    let processed: string[] = [];
    if (existsSync(marker)) processed = JSON.parse(readFileSync(marker, 'utf8'));
    for (const item of history.filter((entry) => entry.mode === 'code' && !entry.summary.scriptId).reverse()) {
      if (processed.includes(item.id)) continue;
      const script = join(item.output, 'script.py');
      if (existsSync(script) && statSync(script).size <= 800000) {
        const code = new TextDecoder('utf-8', { fatal: true }).decode(readFileSync(script));
        const record = this.recordRun(code, `${item.name} · 历史脚本`, item.date);
        const status = item.summary.status;
        this.finishRun(record.id, ['success', 'error', 'stopped', 'timeout'].includes(String(status)) ? status as CodeResult['status'] : 'success');
      }
      processed.push(item.id);
      this.atomicWrite(marker, JSON.stringify(processed));
    }
  }
}
