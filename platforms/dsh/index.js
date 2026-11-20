/**
 * dsh-weekly-report — DeepSeek Harness (DSH) Cordis bundle wrapper.
 *
 * Registers the single-source `weekly-report` skill into the host `skills` service
 * so DSH discovers its routing description on demand without duplicating scripts or docs.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const PLUGIN_ID = 'dsh-weekly-report';
const SKILL_NAME = 'weekly-report';

function stripQuotes(value) {
  const text = value.trim();
  if (text.length >= 2) {
    const first = text[0];
    const last = text.at(-1);
    if ((first === '"' && last === '"') || (first === "'" && last === "'")) {
      return text.slice(1, -1);
    }
  }
  return text;
}

function parseFrontmatter(text, file) {
  const match = /^---\r?\n([\s\S]*?)\r?\n---\r?\n+([\s\S]+)$/.exec(text);
  if (!match) {
    throw new Error(`${PLUGIN_ID}: ${file} requires YAML frontmatter and body`);
  }
  const fields = new Map();
  for (const line of match[1].split(/\r?\n/)) {
    if (!line.trim() || line.trimStart().startsWith('#') || /^\s/.test(line)) continue;
    const field = /^([A-Za-z][A-Za-z0-9_-]*):\s*(.*)$/.exec(line);
    if (!field || field[2] === '') continue;
    fields.set(field[1], stripQuotes(field[2]));
  }
  return { fields, content: match[2].trim() };
}

function loadBundledSkill() {
  const directory = path.join(HERE, 'skills', SKILL_NAME);
  const file = path.join(directory, 'SKILL.md');
  const { fields, content } = parseFrontmatter(fs.readFileSync(file, 'utf8'), file);
  const declaredName = fields.get('name');
  const description = fields.get('description');
  if (declaredName !== SKILL_NAME) {
    throw new Error(`${PLUGIN_ID}: ${file} declares "${declaredName}", expected "${SKILL_NAME}"`);
  }
  if (!description?.trim()) {
    throw new Error(`${PLUGIN_ID}: ${file} requires a non-empty description`);
  }
  return {
    name: SKILL_NAME,
    description: description.trim(),
    source: 'bundled',
    path: file,
    resourceBase: { kind: 'directory', path: directory },
    invocation: {
      modelInvocable: true,
      userInvocable: true,
    },
    content,
  };
}

export function apply(ctx, config = {}) {
  if (config?.enabled === false) return () => {};
  const skill = loadBundledSkill();

  const fiber = ctx.inject(['skills'], (scope) => {
    const service = scope.get('skills');
    if (typeof service?.register !== 'function') {
      throw new Error(`${PLUGIN_ID}: host skills service does not expose register()`);
    }
    const dispose = service.register(skill);
    if (typeof dispose !== 'function') {
      throw new Error(`${PLUGIN_ID}: skills.register("${skill.name}") must return a disposer`);
    }
    return () => dispose();
  });

  return () => {
    if (fiber && typeof fiber.dispose === 'function') fiber.dispose();
    else if (typeof fiber === 'function') fiber();
  };
}

export const name = PLUGIN_ID;
