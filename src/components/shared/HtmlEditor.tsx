/**
 * HtmlEditor — WYSIWYG HTML editor using TinyMCE 6
 * Bundled with SPFx (no CDN — avoids CSP restrictions on SharePoint Online).
 * React.createElement throughout, ES5-safe, Segoe UI, Forest Teal theme.
 *
 * Adapted from DWx LearnIQ Course Builder implementation.
 */
import * as React from 'react';

/* TinyMCE loaded lazily on first editor mount — not at module parse time.
   This avoids adding ~200-300KB to every webpart's initial bundle. */
/* eslint-disable @typescript-eslint/no-var-requires */
var tinymce: any = null;
var tinymceLoaded = false;

function loadTinyMCE(): any {
  if (tinymceLoaded) return tinymce;
  try {
    tinymce = require('tinymce');
    require('tinymce/themes/silver');
    require('tinymce/icons/default');
    require('tinymce/models/dom');
    require('tinymce/skins/ui/oxide/skin.js');
    require('tinymce/skins/ui/oxide/content.js');
    require('tinymce/skins/content/default/content.js');
    require('tinymce/plugins/advlist');
    require('tinymce/plugins/autolink');
    require('tinymce/plugins/lists');
    require('tinymce/plugins/link');
    require('tinymce/plugins/image');
    require('tinymce/plugins/charmap');
    require('tinymce/plugins/preview');
    require('tinymce/plugins/searchreplace');
    require('tinymce/plugins/visualblocks');
    require('tinymce/plugins/code');
    require('tinymce/plugins/fullscreen');
    require('tinymce/plugins/insertdatetime');
    require('tinymce/plugins/media');
    require('tinymce/plugins/table');
    require('tinymce/plugins/wordcount');
    require('tinymce/plugins/codesample');
    tinymceLoaded = true;
  } catch (e) {
    /* TinyMCE not available — will show fallback */
  }
  return tinymce;
}

export interface IHtmlEditorProps {
  value?: string;
  onChange?: (html: string) => void;
  onSave?: (html: string) => void;
  placeholder?: string;
  height?: number;
  readOnly?: boolean;
}

var FONT = "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif";

/* Inline skin CSS to avoid CSP issues */
var SKIN_CSS = '.tox.tox-tinymce{visibility:visible!important;width:100%!important;min-width:0!important}.tox .tox-editor-header{z-index:1}.tox .tox-toolbar__primary{background:#F8FAFC!important;border-bottom:1px solid #E2E8F0!important;flex-wrap:wrap!important}.tox .tox-edit-area__iframe{background:#fff}.tox .tox-statusbar{border-top:1px solid #E2E8F0!important;background:#F8FAFC!important}';

export var HtmlEditor: React.FC<IHtmlEditorProps> = function (props) {
  var editorContainerRef = React.useRef<HTMLDivElement>(null);
  var editorInstanceRef = React.useRef<any>(null);

  var readyState = React.useState(false);
  var isReady = readyState[0];
  var setIsReady = readyState[1];

  var previewState = React.useState(false);
  var showPreview = previewState[0];
  var setShowPreview = previewState[1];

  var htmlState = React.useState(props.value || '');
  var currentHtml = htmlState[0];
  var setCurrentHtml = htmlState[1];

  var errorState = React.useState('');
  var editorError = errorState[0];
  var setEditorError = errorState[1];

  React.useEffect(function () {
    // Lazy-load TinyMCE on first editor mount
    var tmce = loadTinyMCE();
    if (!tmce) {
      setEditorError('TinyMCE could not be loaded. Use the plain text editor instead.');
      return;
    }

    var editorId = 'pm-tinymce-' + Math.random().toString(36).substring(2, 9);
    if (editorContainerRef.current) {
      var textarea = document.createElement('textarea');
      textarea.id = editorId;
      textarea.style.visibility = 'hidden';
      editorContainerRef.current.appendChild(textarea);
    }

    if (!document.getElementById('pm-tinymce-skin')) {
      var style = document.createElement('style');
      style.id = 'pm-tinymce-skin';
      style.textContent = SKIN_CSS;
      document.head.appendChild(style);
    }

    tmce.init({
      selector: '#' + editorId,
      height: props.height || 450,
      menubar: true,
      readonly: props.readOnly || false,
      placeholder: props.placeholder || 'Start writing your policy content...',
      skin: 'oxide',
      content_css: 'default',
      plugins: 'advlist autolink lists link image charmap preview searchreplace visualblocks code fullscreen insertdatetime media table wordcount codesample',
      toolbar: 'undo redo | blocks fontfamily fontsize | bold italic underline strikethrough | forecolor backcolor | alignleft aligncenter alignright alignjustify | bullist numlist outdent indent | table link image media codesample | blockquote | removeformat | fullscreen code',
      toolbar_mode: 'sliding',
      content_style: [
        'body { font-family: Segoe UI, -apple-system, sans-serif; font-size: 15px; color: #0f172a; line-height: 1.7; padding: 16px; }',
        'h1 { font-size: 28px; font-weight: 700; margin-bottom: 12px; }',
        'h2 { font-size: 22px; font-weight: 700; color: #0f766e; margin: 24px 0 8px; }',
        'h3 { font-size: 18px; font-weight: 600; margin: 20px 0 8px; }',
        'blockquote { border-left: 4px solid #99f6e4; padding: 12px 20px; background: #f0fdfa; border-radius: 0 8px 8px 0; margin: 16px 0; }',
        'pre { background: #f1f5f9; padding: 16px; border-radius: 8px; }',
        'code { background: #f1f5f9; padding: 2px 6px; border-radius: 4px; font-size: 14px; }',
        'img { max-width: 100%; border-radius: 8px; }',
        'table { border-collapse: collapse; width: 100%; }',
        'th, td { border: 1px solid #e2e8f0; padding: 10px 14px; }',
        'th { background: #f8fafc; font-weight: 600; }',
        'a { color: #0d9488; }'
      ].join(' '),
      promotion: false,
      branding: false,
      resize: true,
      statusbar: true,
      setup: function (editor: any) {
        editor.on('init', function () {
          if (props.value) {
            editor.setContent(props.value);
          }
          var container = editor.getContainer();
          if (container) {
            container.style.visibility = 'visible';
          }
          setIsReady(true);
        });
        editor.on('change keyup', function () {
          var html = editor.getContent();
          setCurrentHtml(html);
          if (props.onChange) { props.onChange(html); }
        });
        editorInstanceRef.current = editor;
      }
    });

    return function () {
      if (editorInstanceRef.current) {
        try { editorInstanceRef.current.remove(); } catch (_e) { /* silent */ }
        editorInstanceRef.current = null;
      }
    };
  }, []);

  function handleSave(): void {
    var html = currentHtml;
    if (editorInstanceRef.current) {
      html = editorInstanceRef.current.getContent();
    }
    var fullHtml = '<!DOCTYPE html>\n<html lang="en">\n<head>\n<meta charset="UTF-8">\n<meta name="viewport" content="width=device-width, initial-scale=1.0">\n<style>\nbody { font-family: \'Segoe UI\', -apple-system, sans-serif; padding: 32px; max-width: 900px; margin: 0 auto; color: #0f172a; line-height: 1.7; }\nh1 { font-size: 28px; font-weight: 700; margin-bottom: 12px; }\nh2 { font-size: 22px; font-weight: 700; color: #0f766e; margin: 24px 0 8px; }\nh3 { font-size: 18px; font-weight: 600; margin: 20px 0 8px; }\np { margin-bottom: 16px; }\nimg { max-width: 100%; border-radius: 12px; margin: 16px 0; }\nblockquote { border-left: 4px solid #99f6e4; padding: 12px 20px; background: #f0fdfa; border-radius: 0 8px 8px 0; margin: 16px 0; color: #0f766e; }\npre { background: #f1f5f9; padding: 16px; border-radius: 8px; overflow-x: auto; }\ncode { background: #f1f5f9; padding: 2px 6px; border-radius: 4px; font-size: 14px; }\ntable { width: 100%; border-collapse: collapse; margin: 16px 0; }\nth, td { padding: 10px 14px; border: 1px solid #e2e8f0; text-align: left; }\nth { background: #f8fafc; font-weight: 600; }\na { color: #0d9488; }\nul, ol { margin-bottom: 16px; padding-left: 24px; }\nli { margin-bottom: 6px; }\n</style>\n</head>\n<body>\n' + html + '\n</body>\n</html>';
    if (props.onSave) { props.onSave(fullHtml); }
  }

  return React.createElement('div', { style: { fontFamily: FONT } },
    /* Editor/Preview toggle + Save button */
    React.createElement('div', {
      style: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }
    },
      React.createElement('div', { style: { display: 'flex', gap: 8 } },
        React.createElement('button', {
          style: { padding: '6px 14px', borderRadius: 4, border: '1px solid ' + (showPreview ? '#e2e8f0' : '#0d9488'), background: showPreview ? '#fff' : '#f0fdfa', color: showPreview ? '#64748b' : '#0d9488', fontSize: 12, fontWeight: 600, cursor: 'pointer', fontFamily: FONT },
          onClick: function () { setShowPreview(false); }
        }, 'Editor'),
        React.createElement('button', {
          style: { padding: '6px 14px', borderRadius: 4, border: '1px solid ' + (showPreview ? '#0d9488' : '#e2e8f0'), background: showPreview ? '#f0fdfa' : '#fff', color: showPreview ? '#0d9488' : '#64748b', fontSize: 12, fontWeight: 600, cursor: 'pointer', fontFamily: FONT },
          onClick: function () {
            if (editorInstanceRef.current) { setCurrentHtml(editorInstanceRef.current.getContent()); }
            setShowPreview(true);
          }
        }, 'Preview')
      ),
      props.onSave ? React.createElement('button', {
        style: { padding: '6px 16px', borderRadius: 4, border: 'none', background: '#0d9488', color: '#fff', fontSize: 12, fontWeight: 600, cursor: 'pointer', fontFamily: FONT, display: 'flex', alignItems: 'center', gap: 6 },
        onClick: handleSave
      },
        React.createElement('svg', { width: 12, height: 12, viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', strokeWidth: 2 },
          React.createElement('path', { d: 'M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z' }),
          React.createElement('polyline', { points: '17 21 17 13 7 13 7 21' })
        ),
        'Save HTML'
      ) : null
    ),
    editorError ? React.createElement('div', {
      style: { padding: '12px 16px', background: '#fef3c7', border: '1px solid #fde68a', borderRadius: 4, fontSize: 13, color: '#92400e', marginBottom: 8 }
    }, editorError) : null,
    !isReady && !showPreview && !editorError ? React.createElement('div', {
      style: { border: '1px solid #e2e8f0', borderRadius: 10, padding: 48, textAlign: 'center' as const, background: '#fff', minHeight: props.height || 450, display: 'flex', flexDirection: 'column' as const, alignItems: 'center', justifyContent: 'center' }
    },
      React.createElement('div', { style: { width: 32, height: 32, border: '3px solid #e2e8f0', borderTopColor: '#0d9488', borderRadius: '50%', animation: 'dwx-spin 0.8s linear infinite', marginBottom: 12 } }),
      React.createElement('p', { style: { fontSize: 13, color: '#94a3b8' } }, 'Loading editor...')
    ) : null,
    React.createElement('div', { ref: editorContainerRef, style: { display: !showPreview && isReady ? 'block' : 'none' } }),
    showPreview ? React.createElement('div', {
      style: { border: '1px solid #e2e8f0', borderRadius: 10, padding: 24, minHeight: props.height || 450, background: '#fff', fontFamily: FONT, lineHeight: 1.7, fontSize: 15, color: '#0f172a', overflow: 'auto' },
      dangerouslySetInnerHTML: { __html: currentHtml }
    }) : null,
    React.createElement('div', {
      style: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: 8, fontSize: 11, color: '#94a3b8' }
    },
      React.createElement('span', null, 'TinyMCE — tables, images, media, code blocks, and more'),
      React.createElement('span', null, currentHtml.length > 0 ? currentHtml.length + ' chars' : '')
    )
  );
};
