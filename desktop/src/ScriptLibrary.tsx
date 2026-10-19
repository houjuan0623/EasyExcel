// Copyright (c) 2026 AbeLin · MIT License
import { useEffect, useMemo, useState } from 'react';
import { ArchiveRestore, BookMarked, Download, FileCode2, LoaderCircle, Play, Save, Search, Star, Trash2, Upload, X } from 'lucide-react';
import type { ScriptInfo, ScriptLibrarySnapshot, ScriptRecord } from '../shared/types';

const api = window.desktop!;
const labels: Record<ScriptInfo['status'], string> = { never: '尚未运行', running: '运行中 / 未结束', success: '运行成功', error: '运行出错', stopped: '已停止', timeout: '已超时' };
const date = (value: string | null) => value ? new Date(value).toLocaleString('zh-CN', { month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' }) : '未运行';
const message = (error: unknown) => error instanceof Error ? error.message.replace(/^Error invoking remote method '[^']+': (?:Error: )?/, '') : String(error);

export default function ScriptLibrary({ onClose, onLoad, canRun, initialId }: {
  onClose: () => void; onLoad: (script: ScriptRecord, run: boolean) => Promise<void>; canRun: boolean; initialId?: string;
}) {
  const [snapshot, setSnapshot] = useState<ScriptLibrarySnapshot | null>(null);
  const [query, setQuery] = useState('');
  const [scope, setScope] = useState<'all' | 'favorite' | 'trash'>('all');
  const [category, setCategory] = useState('');
  const [selected, setSelected] = useState(initialId || '');
  const [record, setRecord] = useState<ScriptRecord | null>(null);
  const [name, setName] = useState('');
  const [group, setGroup] = useState('');
  const [notes, setNotes] = useState('');
  const [busy, setBusy] = useState(false);
  const [reading, setReading] = useState(false);
  const [error, setError] = useState('');
  const [info, setInfo] = useState('');

  async function refresh() { setSnapshot(await api.scripts()); }
  useEffect(() => { void refresh().catch((e) => setError(message(e))); }, []);
  const categories = useMemo(() => [...new Set(snapshot?.items.map((item) => item.category))].sort(), [snapshot]);
  const items = useMemo(() => (snapshot?.items || []).filter((item) => {
    return (scope === 'trash' ? item.archived : !item.archived && (scope !== 'favorite' || item.favorite)) &&
      (!category || item.category === category) && `${item.name} ${item.category} ${item.notes}`.toLowerCase().includes(query.toLowerCase());
  }), [snapshot, scope, category, query]);
  useEffect(() => { if (snapshot && !items.some((item) => item.id === selected)) setSelected(items[0]?.id || ''); }, [items, snapshot, selected]);
  useEffect(() => {
    let current = true;
    setRecord(null);
    if (!selected) { setReading(false); return; }
    setReading(true);
    api.script(selected).then((value) => {
      if (current) { setRecord(value); setName(value.name); setGroup(value.category); setNotes(value.notes); }
    }).catch((e) => { if (current) setError(message(e)); }).finally(() => { if (current) setReading(false); });
    return () => { current = false; };
  }, [selected]);

  async function action(task: () => Promise<void>) {
    if (busy) return;
    setBusy(true); setError(''); setInfo('');
    try { await task(); } catch (e) { setError(message(e)); } finally { setBusy(false); }
  }
  async function saveMetadata() {
    if (!record) return null;
    const updated = await api.updateScript(record.id, { name, category: group, notes });
    setRecord(updated); setName(updated.name); setGroup(updated.category); setNotes(updated.notes);
    await refresh();
    return updated;
  }
  async function changeFlag(changes: { favorite?: boolean; archived?: boolean }) {
    if (!record) return;
    const updated = await api.updateScript(record.id, changes);
    setRecord(updated);
    await refresh();
  }

  return <div className="modal-backdrop" onClick={() => { if (!busy) onClose(); }}>
    <section className="script-library" role="dialog" aria-modal="true" aria-label="本地脚本库" onClick={(e) => e.stopPropagation()}>
      <div className="library-heading"><div><h2><BookMarked size={21} />本地脚本库</h2><p>运行过的代码自动收录，同一份代码累计次数。加载后使用当前勾选的 Sheet。</p></div><button className="icon-button" aria-label="关闭脚本库" disabled={busy} onClick={onClose}><X size={19} /></button></div>
      <div className="library-toolbar"><label className="search-field"><Search size={15} /><input aria-label="搜索脚本" placeholder="搜索名称、分类、说明…" value={query} onChange={(e) => setQuery(e.target.value)} disabled={busy} /></label><select aria-label="筛选脚本分类" value={category} onChange={(e) => setCategory(e.target.value)} disabled={busy}><option value="">全部分类</option>{categories.map((value) => <option key={value}>{value}</option>)}</select><button className="button small" disabled={busy} onClick={() => void action(async () => { const path = await api.exportScripts(); if (path) setInfo(`备份已保存：${path}`); })}><Download size={14} />备份</button><button className="button small" disabled={busy} onClick={() => void action(async () => { const result = await api.importScripts(); if (result) { await refresh(); setInfo(`已导入 ${result.added} 份，合并重复记录 ${result.merged} 份。`); } })}><Upload size={14} />导入备份</button></div>
      <div className="library-scopes">{([['all', '全部脚本'], ['favorite', '我的收藏'], ['trash', '回收站']] as const).map(([value, title]) => <button key={value} className={scope === value ? 'active' : ''} disabled={busy} onClick={() => setScope(value)}>{title}<span>{snapshot?.items.filter((item) => value === 'trash' ? item.archived : !item.archived && (value !== 'favorite' || item.favorite)).length || 0}</span></button>)}</div>
      {error && <div className="library-notice error" role="alert">{error}</div>}
      {info && <div className="library-notice" role="status">{info}</div>}
      {!!snapshot?.warnings.length && <div className="library-notice error">{snapshot.warnings.join('\n')}</div>}
      <div className="library-body">
        <div className="library-list" aria-label="脚本列表">{items.map((item) => <button key={item.id} className={`library-item ${selected === item.id ? 'selected' : ''}`} disabled={busy} onClick={() => setSelected(item.id)}><span className="library-item-title"><FileCode2 size={16} /><strong>{item.name}</strong>{item.favorite && <Star size={13} fill="currentColor" />}</span><span className="library-item-meta">{item.category} · {item.runCount} 次运行 · {item.lines} 行</span><span className={`library-item-status ${item.status}`}>{labels[item.status]}<span>{date(item.lastRunAt || item.updatedAt)}</span></span></button>)}{!items.length && <div className="library-empty"><BookMarked size={30} /><p>{snapshot ? scope === 'trash' ? '回收站为空' : query || category ? '没有匹配的脚本' : '运行一次代码，或将编辑内容存入脚本库。' : '正在读取脚本库…'}</p></div>}</div>
        <div className="library-detail">{reading ? <div className="library-empty"><LoaderCircle className="spin" size={25} /></div> : record ? <>
          <div className="library-detail-fields"><label>脚本名称<input aria-label="脚本名称" value={name} maxLength={120} disabled={busy || record.archived} onChange={(e) => setName(e.target.value)} /></label><label>分类<input aria-label="脚本分类" list="script-categories" value={group} maxLength={40} disabled={busy || record.archived} onChange={(e) => setGroup(e.target.value)} /><datalist id="script-categories">{[...new Set(['财务核对', '报表汇总', '现金流', '图表分析', '草稿', ...categories])].map((value) => <option key={value} value={value} />)}</datalist></label></div>
          <label className="library-description">使用说明<textarea aria-label="脚本使用说明" value={notes} maxLength={2000} disabled={busy || record.archived} onChange={(e) => setNotes(e.target.value)} placeholder="适用场景、需要的字段、参数与注意事项" /></label>
          <div className="library-detail-actions">{!record.archived ? <><button className="button small" disabled={busy} onClick={() => void action(async () => { await saveMetadata(); setInfo('脚本信息已保存。'); })}><Save size={14} />保存信息</button><button className={`button small ${record.favorite ? 'favorited' : ''}`} disabled={busy} onClick={() => void action(() => changeFlag({ favorite: !record.favorite }))}><Star size={14} fill={record.favorite ? 'currentColor' : 'none'} />{record.favorite ? '取消收藏' : '收藏脚本'}</button><button className="text-button subtle" disabled={busy} onClick={() => void action(() => changeFlag({ archived: true }))}><Trash2 size={14} />移入回收站</button></> : <button className="button small" disabled={busy} onClick={() => void action(() => changeFlag({ archived: false }))}><ArchiveRestore size={14} />恢复脚本</button>}</div>
          <div className="library-preview-label"><span>代码预览 · {record.lines} 行</span><span>创建于 {date(record.createdAt)}</span></div><pre className="library-code">{record.code}</pre>
          <div className="library-load-actions"><p>载入前，当前未收录的编辑内容会自动暂存为草稿。</p><div><button className="button" disabled={busy || record.archived} onClick={() => void action(async () => { const saved = await saveMetadata(); if (saved) await onLoad(saved, false); })}>载入编辑器</button><button className="button primary" disabled={busy || record.archived || !canRun} title={canRun ? '使用当前勾选 Sheet 运行' : '先在工作台选择表格与 Sheet'} onClick={() => void action(async () => { const saved = await saveMetadata(); if (saved) await onLoad(saved, true); })}><Play size={14} />载入并运行</button></div></div>
        </> : <div className="library-empty"><FileCode2 size={30} /><p>选择一份脚本查看内容</p></div>}</div>
      </div>
      <div className="library-footer"><span title={snapshot?.path}>保存在本机用户目录 · 更新程序后保留 · 换电脑可备份迁移</span><span>{items.length} 份脚本</span></div>
    </section>
  </div>;
}
