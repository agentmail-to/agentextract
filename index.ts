// agentextract — package entry point.
//
// Two peer capabilities, surfaced side by side: email BODY extraction (strip quoted history and
// signature/boilerplate) and email ATTACHMENT text extraction (pdf/docx/xlsx/...). Neither is
// primary — the top-level import gives you both, and each also stands alone on its own subpath
// ('agentextract/body' and 'agentextract/attachment'). Attachment handlers lazy-load their heavy
// parsers, so pulling in the whole package costs nothing until extractAttachment is actually called.

export * from './body'

export { extractAttachment, detectRoute } from './attachment'
export type {
    AttachmentInput,
    ExtractionResult,
    ExtractionStatus,
    HandlerKind,
    RoutedBy,
} from './attachment'
