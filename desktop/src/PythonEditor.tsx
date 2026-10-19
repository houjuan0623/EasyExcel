// Copyright (c) 2026 AbeLin · MIT License
import { useEffect, useRef } from 'react';
import { basicSetup } from 'codemirror';
import { python } from '@codemirror/lang-python';
import { Compartment, EditorState, Prec } from '@codemirror/state';
import { EditorView, keymap } from '@codemirror/view';
import { indentWithTab } from '@codemirror/commands';

export default function PythonEditor({ value, disabled, onChange, onRun }: {
  value: string; disabled: boolean; onChange: (value: string) => void; onRun: () => void;
}) {
  const parent = useRef<HTMLDivElement>(null);
  const editor = useRef<EditorView | null>(null);
  const editable = useRef(new Compartment());
  const callbacks = useRef({ onChange, onRun });
  callbacks.current = { onChange, onRun };
  useEffect(() => {
    const view = new EditorView({ parent: parent.current!, state: EditorState.create({ doc: value, extensions: [
      basicSetup, python(), EditorState.tabSize.of(4), editable.current.of(EditorState.readOnly.of(disabled)),
      Prec.highest(keymap.of([{ key: 'Mod-Enter', run: () => { callbacks.current.onRun(); return true; } }, indentWithTab])),
      EditorView.contentAttributes.of({ 'aria-label': 'Python 代码', role: 'textbox', 'aria-multiline': 'true', spellcheck: 'false' }),
      EditorView.updateListener.of((update) => { if (update.docChanged) callbacks.current.onChange(update.state.doc.toString()); }),
      EditorView.theme({ '&': { height: '370px' }, '.cm-scroller': { overflow: 'auto', fontFamily: '"Cascadia Code", Consolas, "Microsoft YaHei", monospace', fontSize: '12px', lineHeight: '1.85' }, '.cm-content': { padding: '12px 0' }, '.cm-gutters': { backgroundColor: '#f4f7f5', color: '#93a89a', borderRight: '1px solid #e8eee9' }, '.cm-activeLine': { backgroundColor: '#eaf5ee66' }, '&.cm-focused': { outline: 'none' } }),
    ] }) });
    editor.current = view;
    return () => { view.destroy(); editor.current = null; };
  }, []); // One editor per mounted workspace; updates preserve selection and undo history.
  useEffect(() => { const view = editor.current; if (view && view.state.doc.toString() !== value) view.dispatch({ changes: { from: 0, to: view.state.doc.length, insert: value } }); }, [value]);
  useEffect(() => { editor.current?.dispatch({ effects: editable.current.reconfigure(EditorState.readOnly.of(disabled)) }); }, [disabled]);
  return <div className="python-editor" ref={parent} />;
}
