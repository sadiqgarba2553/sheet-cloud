import { useEffect, useRef, ReactNode } from 'react';
import { X } from 'lucide-react';
export function Modal({ title, children, close }: { title: string; children: ReactNode; close: () => void }) {
  const ref = useRef<HTMLDialogElement>(null);
  useEffect(() => { const dialog = ref.current!; dialog.showModal(); return () => dialog.close(); }, []);
  return <dialog ref={ref} className="modal" onCancel={close} onClick={e => { if (e.target === ref.current) close(); }}><header><h2>{title}</h2><button className="icon" aria-label="Close dialog" onClick={close}><X size={18} /></button></header><div className="modal-body">{children}</div></dialog>;
}
