/**
 * Excel2PPT Studio · DSH Client half.
 *
 * 在输入框下方的 dock 放一个入口按钮，点击打开本地 Studio 页面。
 * 不嵌入 iframe：Harness 页面与本地应用主题互相独立，外部标签页体验更稳。
 */
window.__ModuleLoader__.load({
  id: '@local/excel2ppt-studio',
  factory(require) {
    const React = require('react');
    const h = React.createElement;

    const STUDIO_URL = 'http://127.0.0.1:8765';

    function Launcher() {
      return h(
        'button',
        {
          type: 'button',
          title: '打开 Excel2PPT Studio（本地四步向导）',
          'aria-label': '打开 Excel2PPT Studio',
          onClick: () => window.open(STUDIO_URL, '_blank', 'noopener'),
          style: {
            display: 'inline-flex',
            alignItems: 'center',
            gap: '6px',
            height: '30px',
            padding: '0 10px',
            marginLeft: '6px',
            borderRadius: '8px',
            border: '1px solid var(--dsw-alias-border, #d5dbe5)',
            background: 'var(--dsw-alias-surface, #ffffff)',
            color: 'var(--dsw-alias-text, #1f2937)',
            cursor: 'pointer',
            font: 'inherit',
            fontSize: '12px',
            lineHeight: 1,
            whiteSpace: 'nowrap',
          },
        },
        h('span', {
          style: {
            display: 'inline-block',
            width: '8px',
            height: '8px',
            borderRadius: '2px',
            background: '#2563eb',
          },
        }),
        'Excel2PPT',
      );
    }

    return {
      inject: ['slots'],
      apply(ctx) {
        ctx.slots.inject('conversation.composer.dock', () =>
          ctx.slots.register(
            {
              name: 'conversation.composer.dock',
              id: 'excel2ppt-launcher',
              order: 40,
            },
            Launcher,
          ),
        );
      },
    };
  },
});
