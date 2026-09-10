// The generator the adversarial fuzz tests draw their mutations from. Deterministic, so a failing case
// is reproducible from its seed alone rather than being a flake someone has to reproduce by luck. Each
// fuzzer keeps its own seeds.

/** xorshift32: the next unsigned 32-bit value on each call, from a nonzero state derived from `seed`. */
export function xorshift32(seed: number): () => number {
  let state = seed | 0 || 1;
  return () => {
    state ^= state << 13;
    state ^= state >>> 17;
    state ^= state << 5;
    return state >>> 0;
  };
}
