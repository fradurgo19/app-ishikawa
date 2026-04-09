import React from 'react';
import { Navigate, useLocation, useNavigate } from 'react-router-dom';
import { ArrowLeft, ExternalLink, FileText, Paperclip } from 'lucide-react';
import { Button } from '../atoms/Button';
import type { Attachment, FishboneDiagramDetailPayload } from '../types';
import { looksLikeNavigableUrl, openUrlInNewBrowserTab } from '../utils/openExternalInNewTab';

function isFromDataTableState(state: unknown): boolean {
  return (
    typeof state === 'object' &&
    state !== null &&
    'fromDataTable' in state &&
    (state as { fromDataTable?: boolean }).fromDataTable === true
  );
}

function isDiagramDetailPayload(value: unknown): value is FishboneDiagramDetailPayload {
  if (!value || typeof value !== 'object') {
    return false;
  }
  const o = value as { kind?: string; recordId?: string; allAttachments?: unknown };
  if (typeof o.recordId !== 'string' || !Array.isArray(o.allAttachments)) {
    return false;
  }
  if (o.kind === 'resource') {
    return typeof (value as { resourceText?: string }).resourceText === 'string';
  }
  if (o.kind === 'attachments') {
    return typeof (value as { focusAttachmentId?: string }).focusAttachmentId === 'string';
  }
  return false;
}

interface FishboneDiagramLocationState {
  fromDataTable?: boolean;
  returnTo?: string;
  diagramDetail?: unknown;
}

export const FishboneDetailPage: React.FC = () => {
  const navigate = useNavigate();
  const location = useLocation();
  const state = location.state as FishboneDiagramLocationState | null;

  if (!state || !isFromDataTableState(state)) {
    return <Navigate to="/data-table" replace />;
  }

  const navState = state;

  const detail = navState.diagramDetail;
  if (!isDiagramDetailPayload(detail)) {
    const fallback =
      typeof navState.returnTo === 'string' && navState.returnTo.startsWith('/')
        ? navState.returnTo
        : '/fishbone';
    return <Navigate to={fallback} replace state={{ fromDataTable: true }} />;
  }

  const returnTo =
    typeof navState.returnTo === 'string' && navState.returnTo.startsWith('/')
      ? navState.returnTo
      : '/fishbone';

  const goBack = () => {
    navigate(returnTo, { state: { fromDataTable: true } });
  };

  return (
    <div className="min-h-screen w-full bg-gray-50">
      <div className="w-full px-4 py-8 sm:px-6 lg:px-8">
        <div className="mx-auto max-w-2xl">
          <Button variant="ghost" icon={ArrowLeft} onClick={goBack} className="mb-6">
            Volver al diagrama
          </Button>

          {detail.kind === 'resource' ? (
            <ResourceDetailView detail={detail} />
          ) : (
            <AttachmentsDetailView detail={detail} />
          )}
        </div>
      </div>
    </div>
  );
};

function ResourceDetailView({
  detail,
}: Readonly<{ detail: Extract<FishboneDiagramDetailPayload, { kind: 'resource' }> }>) {
  const text = detail.resourceText.trim();
  const href = looksLikeNavigableUrl(text) ? text : undefined;

  let resourceBlock: React.ReactNode;
  if (!text) {
    resourceBlock = <p className="text-gray-500">No hay texto de recurso en este registro.</p>;
  } else if (href) {
    resourceBlock = (
      <button
        type="button"
        onClick={() => {
          openUrlInNewBrowserTab(href);
        }}
        className="break-words text-left text-base text-blue-600 underline hover:text-blue-800"
      >
        {text}
      </button>
    );
  } else {
    resourceBlock = (
      <p className="whitespace-pre-wrap break-words text-base text-gray-800">{detail.resourceText}</p>
    );
  }

  return (
    <div className="rounded-lg border border-gray-200 bg-white p-6 shadow-md">
      <div className="mb-4 flex items-center gap-2 text-gray-900">
        <FileText className="h-6 w-6 shrink-0 text-teal-600" aria-hidden />
        <h1 className="text-xl font-bold">Recurso</h1>
      </div>
      <p className="mb-2 text-xs text-gray-500">Registro: {detail.recordId}</p>
      {resourceBlock}
      <AttachmentsListSection attachments={detail.allAttachments} title="Adjuntos del mismo registro" />
    </div>
  );
}

function AttachmentsDetailView({
  detail,
}: Readonly<{ detail: Extract<FishboneDiagramDetailPayload, { kind: 'attachments' }> }>) {
  return (
    <div className="rounded-lg border border-gray-200 bg-white p-6 shadow-md">
      <div className="mb-4 flex items-center gap-2 text-gray-900">
        <Paperclip className="h-6 w-6 shrink-0 text-pink-600" aria-hidden />
        <h1 className="text-xl font-bold">Adjuntos</h1>
      </div>
      <p className="mb-4 text-xs text-gray-500">Registro: {detail.recordId}</p>
      <AttachmentsListSection
        attachments={detail.allAttachments}
        highlightId={detail.focusAttachmentId}
        title="Archivos"
      />
    </div>
  );
}

function AttachmentsListSection({
  attachments,
  title,
  highlightId,
}: Readonly<{ attachments: Attachment[]; title: string; highlightId?: string }>) {
  if (attachments.length === 0) {
    return (
      <div className="mt-8 border-t border-gray-100 pt-6">
        <h2 className="mb-2 text-sm font-semibold text-gray-800">{title}</h2>
        <p className="text-sm text-gray-500">No hay adjuntos en este registro.</p>
      </div>
    );
  }

  return (
    <div className="mt-8 border-t border-gray-100 pt-6">
      <h2 className="mb-3 text-sm font-semibold text-gray-800">{title}</h2>
      <ul className="m-0 list-none space-y-2 p-0">
        {attachments.map((att, index) => {
          const href = att.url?.trim();
          const label = att.name?.trim() || 'Adjunto';
          const key = `${att.id}-${index}`;
          const isHighlight = highlightId !== undefined && att.id === highlightId;
          const itemClass = isHighlight
            ? 'rounded-md border border-pink-200 bg-pink-50/80 px-3 py-2'
            : 'rounded-md border border-transparent px-3 py-2';

          if (!href) {
            return (
              <li key={key} className={itemClass}>
                <span className="text-sm text-gray-700" title={label}>
                  {label}
                </span>
                <span className="ml-2 text-xs text-gray-400">(sin enlace)</span>
              </li>
            );
          }

          return (
            <li key={key} className={itemClass}>
              <button
                type="button"
                onClick={() => {
                  openUrlInNewBrowserTab(href);
                }}
                className="inline-flex max-w-full items-center gap-2 text-left text-sm text-blue-600 hover:underline"
              >
                <ExternalLink className="h-4 w-4 shrink-0" aria-hidden />
                <span className="break-words" title={label}>
                  {label}
                </span>
              </button>
            </li>
          );
        })}
      </ul>
    </div>
  );
}
