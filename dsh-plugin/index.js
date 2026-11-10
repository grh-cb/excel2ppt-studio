/**
 * Excel2PPT Studio · DSH Host half.
 *
 * 暴露三个模型工具：build / inspect / verify，内部调用 `python -m excel2ppt_studio`。
 * 零 npm 依赖：只用 Node 标准库，避免安装时还要解析 npm 依赖树。
 */
import { spawn } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

export const name = 'excel2ppt';
export const inject = ['tools'];

const DEFAULTS = {
  pythonPath: 'python',
  studioUrl: 'http://127.0.0.1:8765',
  timeoutMs: 180000,
};

function run(cmd, args, timeoutMs) {
  return new Promise((resolve, reject) => {
    const child = spawn(cmd, args, {
      windowsHide: true,
      env: { ...process.env, PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1' },
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    let stdout = '';
    let stderr = '';
    child.stdout.setEncoding('utf8');
    child.stderr.setEncoding('utf8');
    child.stdout.on('data', (chunk) => { stdout += chunk; });
    child.stderr.on('data', (chunk) => { stderr += chunk; });
    const timer = setTimeout(() => {
      child.kill();
      reject(new Error(`命令超时（${Math.round(timeoutMs / 1000)}s）`));
    }, timeoutMs);
    child.on('error', (err) => { clearTimeout(timer); reject(err); });
    child.on('close', (code) => { clearTimeout(timer); resolve({ code, stdout, stderr }); });
  });
}

function lastJson(text) {
  const lines = String(text || '').trim().split(/\r?\n/).filter(Boolean);
  for (let i = lines.length - 1; i >= 0; i -= 1) {
    try { return JSON.parse(lines[i]); } catch { /* 继续向上找 */ }
  }
  return null;
}

function renderJson(_args, value) {
  return [{ type: 'text', text: JSON.stringify(value, null, 2) }];
}

function failure(res) {
  const text = (res.stderr || res.stdout || '').trim();
  return text || `excel2ppt 退出码 ${res.code}`;
}

export function apply(ctx, config = {}) {
  const cfg = { ...DEFAULTS, ...(config || {}) };
  const python = cfg.pythonPath || DEFAULTS.pythonPath;
  const timeoutMs = Number(cfg.timeoutMs) > 0 ? Number(cfg.timeoutMs) : DEFAULTS.timeoutMs;

  ctx.tools.register({
    name: 'excel2ppt_build',
    description:
      '用 Excel2PPT Studio 把 Excel 台账按单页 PPT 模板逐行生成多页幻灯片（克隆母版页、只换内容，排版/颜色/位置零漂移）。需要先有映射 Profile：可在 Studio 界面里建立并导出，或用 excel2ppt_inspect 查看模板元素后手写。返回输出路径、页数、替换处数与告警。',
    parameters: {
      type: 'object',
      additionalProperties: false,
      properties: {
        profile: {
          type: 'object',
          description: '完整 Profile 对象，含 template / excel / sheet / bindings / pagination。',
        },
        profilePath: {
          type: 'string',
          description: '已有 Profile JSON 的绝对路径；与 profile 二选一。',
        },
        excel: { type: 'string', description: '覆盖 Profile 里的 Excel 路径。' },
        template: { type: 'string', description: '覆盖 Profile 里的模板路径。' },
        out: { type: 'string', description: '输出 .pptx 路径；默认与 Excel 同目录。' },
      },
    },
    output: { schema: { type: 'object', additionalProperties: true }, render: renderJson },
    async execute(args) {
      if (!args.profile && !args.profilePath) {
        throw new Error('需要 profile 或 profilePath 之一。');
      }
      const tmp = await mkdtemp(join(tmpdir(), 'e2p-dsh-'));
      try {
        let profilePath = args.profilePath;
        if (!profilePath) {
          profilePath = join(tmp, 'profile.json');
          await writeFile(profilePath, JSON.stringify(args.profile, null, 2), 'utf8');
        }
        const cmd = ['-m', 'excel2ppt_studio', 'build', '--profile', profilePath, '--json'];
        if (args.excel) cmd.push('--excel', args.excel);
        if (args.template) cmd.push('--template', args.template);
        if (args.out) cmd.push('--out', args.out);
        const res = await run(python, cmd, timeoutMs);
        if (res.code !== 0) throw new Error(failure(res));
        const data = lastJson(res.stdout);
        if (!data) throw new Error(`未解析到 JSON 输出：${res.stdout.slice(0, 300)}`);
        return data;
      } finally {
        await rm(tmp, { recursive: true, force: true });
      }
    },
  });

  ctx.tools.register({
    name: 'excel2ppt_inspect',
    description:
      '查看 PPT 模板指定页的可填充元素：类型（文本/表格/图片/图表）、角色、绝对坐标（cm）、文本与槽位名。用于在写映射前确认模板结构。',
    parameters: {
      type: 'object',
      additionalProperties: false,
      properties: {
        template: { type: 'string', description: '模板 .pptx 的绝对路径。' },
        slide: { type: 'integer', description: '母版页索引，从 0 开始；默认 0。' },
      },
      required: ['template'],
    },
    output: { schema: { type: 'object', additionalProperties: true }, render: renderJson },
    async execute(args) {
      const tmp = await mkdtemp(join(tmpdir(), 'e2p-dsh-'));
      try {
        const jsonPath = join(tmp, 'fields.json');
        const cmd = [
          '-m', 'excel2ppt_studio', 'inspect', args.template,
          '--slide', String(args.slide ?? 0), '--json', jsonPath,
        ];
        const res = await run(python, cmd, timeoutMs);
        if (res.code !== 0) throw new Error(failure(res));
        const data = JSON.parse(await readFile(jsonPath, 'utf8'));
        return {
          template: data.path,
          slide_index: data.slide_index,
          slide_count: data.slide_count,
          width_emu: data.width,
          height_emu: data.height,
          fields: data.fields,
          warnings: data.warnings,
        };
      } finally {
        await rm(tmp, { recursive: true, force: true });
      }
    },
  });

  ctx.tools.register({
    name: 'excel2ppt_verify',
    description:
      '零漂移校验：逐页比对输出 PPT 与模板母版页的 shape 集合、位置/尺寸、字体签名，并报告残留占位符。返回 ok、逐条 issues 与页数统计。',
    parameters: {
      type: 'object',
      additionalProperties: false,
      properties: {
        template: { type: 'string', description: '模板 .pptx 的绝对路径。' },
        output: { type: 'string', description: '生成的 .pptx 的绝对路径。' },
        slide: { type: 'integer', description: '模板母版页索引；默认 0。' },
        strictImages: { type: 'boolean', description: '图片尺寸变化也算错误（默认只告警）。' },
      },
      required: ['template', 'output'],
    },
    output: { schema: { type: 'object', additionalProperties: true }, render: renderJson },
    async execute(args) {
      const cmd = [
        '-m', 'excel2ppt_studio', 'verify',
        '--template', args.template, '--output', args.output,
        '--slide', String(args.slide ?? 0), '--json',
      ];
      if (args.strictImages) cmd.push('--strict-images');
      const res = await run(python, cmd, timeoutMs);
      const data = lastJson(res.stdout);
      if (!data) throw new Error(failure(res));
      return data;
    },
  });
}
