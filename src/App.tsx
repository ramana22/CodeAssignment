import React, { useCallback, useMemo, useState } from 'react';

interface RedactionResult {
  redactions: number;
  addedHeader: boolean;
  trackingEnabled: boolean;
}

const patterns = {
  email: {
    regex: /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b/g,
    replacement: '[REDACTED EMAIL]'
  },
  phone: {
    regex: /\b(?:\+?1[-.\s]?)?(?:\(\d{3}\)|\d{3})[-.\s]?\d{3}[-.\s]?\d{4}\b/g,
    replacement: '[REDACTED PHONE]'
  },
  ssn: {
    regex: /\b\d{3}-\d{2}-\d{4}\b/g,
    replacement: '[REDACTED SSN]'
  }
};

const App: React.FC = () => {
  const [status, setStatus] = useState<string>('Ready to redact sensitive information.');
  const [isBusy, setIsBusy] = useState(false);
  const [lastResult, setLastResult] = useState<RedactionResult | null>(null);

  const supportsWordApi15 = useMemo(
    () => (Office.context?.requirements?.isSetSupported('WordApi', '1.5') ?? false),
    []
  );

  const redactDocument = useCallback(async () => {
    setIsBusy(true);
    setStatus('Scanning document for sensitive information...');

    try {
      await Word.run(async (context) => {
        const document = context.document;
        const trackRevisionsAvailable = Office.context.requirements.isSetSupported('WordApi', '1.5');

        if (trackRevisionsAvailable) {
          document.load('trackRevisions');
          await context.sync();
          if (!document.trackRevisions) {
            document.trackRevisions = true;
          }
        } else {
          setStatus('Tracking changes is not available in this host. Redaction will continue without it.');
        }

        const sections = document.sections;
        sections.load('items');
        await context.sync();

        if (sections.items.length === 0) {
          setStatus('The document has no sections to update.');
          return;
        }

        const header = sections.getFirst().getHeader('Primary');
        const headerRange = header.getRange('Whole');
        headerRange.load('text');
        await context.sync();

        let addedHeader = false;
        const headerMarker = 'CONFIDENTIAL DOCUMENT';

        if (!headerRange.text.includes(headerMarker)) {
          header.body.insertParagraph(headerMarker, 'Start');
          addedHeader = true;
        }

        const bodyRange = document.body.getRange();
        bodyRange.load('text');
        await context.sync();

        const contentText = bodyRange.text;
        const matches = new Map<string, string>();

        Object.values(patterns).forEach(({ regex, replacement }) => {
          regex.lastIndex = 0;
          let match: RegExpExecArray | null;
          while ((match = regex.exec(contentText)) !== null) {
            matches.set(match[0], replacement);
          }
        });

        let totalReplacements = 0;

        for (const [foundText, replacement] of matches.entries()) {
          const searchResults = document.body.search(foundText, {
            matchCase: false,
            matchWholeWord: false
          });
          searchResults.load('items');
          await context.sync();

          searchResults.items.forEach((range) => {
            range.insertText(replacement, 'Replace');
            totalReplacements += 1;
          });
        }

        setLastResult({
          redactions: totalReplacements,
          addedHeader,
          trackingEnabled: trackRevisionsAvailable
        });

        setStatus(
          totalReplacements > 0
            ? `Redacted ${totalReplacements} sensitive entr${totalReplacements === 1 ? 'y' : 'ies'}.`
            : 'No sensitive information found in the document.'
        );
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : 'An unexpected error occurred while redacting.';
      setStatus(message);
    } finally {
      setIsBusy(false);
    }
  }, [supportsWordApi15]);

  return (
    <div className="app-shell">
      <header className="hero">
        <p className="eyebrow">Word add-in</p>
        <h1>Document Redaction</h1>
        <p className="lede">
          Replace sensitive information, add a confidentiality header, and keep changes tracked when supported.
        </p>
      </header>

      <div className="card">
        <div className="card-body">
          <div className="status-row">
            <div>
              <p className="status-label">Tracking support</p>
              <p className="status-value">{supportsWordApi15 ? 'Word API 1.5 available' : 'Tracking not available'}</p>
            </div>
            <div>
              <p className="status-label">Last action</p>
              <p className="status-value">{lastResult ? status : 'Pending'}</p>
            </div>
          </div>

          <button className="primary" onClick={redactDocument} disabled={isBusy}>
            {isBusy ? 'Working...' : 'Redact & Track Changes'}
          </button>

          <p className="muted">{status}</p>

          {lastResult && (
            <div className="summary">
              <h2>Redaction summary</h2>
              <ul>
                <li>
                  <span>Entries redacted</span>
                  <strong>{lastResult.redactions}</strong>
                </li>
                <li>
                  <span>Header added</span>
                  <strong>{lastResult.addedHeader ? 'Yes' : 'Already present'}</strong>
                </li>
                <li>
                  <span>Tracking changes</span>
                  <strong>{lastResult.trackingEnabled ? 'Enabled' : 'Unavailable'}</strong>
                </li>
              </ul>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default App;
