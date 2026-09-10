// Comment-thread fixtures shared by the tests that author threads. Not a `.test.ts` file, so
// `node --test` does not try to run it, and `tsconfig.build.json` excludes the suffix from `dist/`.

import type {CommentThread} from './comment-thread.ts';

let nextThread = 0;

/**
 * A one-comment thread anchored at `ref`, with a distinct id per call in the only spelling the format
 * accepts: brace-wrapped, upper-case, `8-4-4-4-12` hex. The authoring path rejects a readable
 * placeholder like `{HEAD-B2}`.
 */
export function threadAt(ref: string): CommentThread {
  return {
    ref,
    resolved: false,
    comments: [
      {
        id: `{${String(++nextThread).padStart(8, '0')}-0000-4000-8000-000000000000}`,
        text: `about ${ref}`,
        mentions: [],
      },
    ],
  };
}
