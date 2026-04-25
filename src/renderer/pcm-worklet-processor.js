// AnchorCast PCM Capture Processor
// Runs in the AudioWorklet thread — captures raw Float32 audio and posts to main thread
// This file must be served as a separate JS file (not inline)

class PcmCaptureProcessor extends AudioWorkletProcessor {
  constructor() {
    super();
    this._buffer = [];
    this._bufferSize = 0;
    // Accumulate ~100ms of audio before sending (reduces IPC overhead)
    // At 16kHz: 1600 samples = 100ms
    // At 44.1kHz: 4410 samples = 100ms
    this._flushSize = 2048; // flush every ~46ms at 44.1kHz (was 4096/~90ms)
  }

  process(inputs) {
    const input = inputs[0];
    if (!input || !input[0]) return true;

    const mono = input[0]; // Float32Array of this chunk (~128 samples)
    this._buffer.push(mono.slice(0)); // copy — avoid detached buffer issues
    this._bufferSize += mono.length;

    if (this._bufferSize >= this._flushSize) {
      // Merge accumulated buffers into one Float32Array
      const merged = new Float32Array(this._bufferSize);
      let offset = 0;
      for (const chunk of this._buffer) {
        merged.set(chunk, offset);
        offset += chunk.length;
      }
      this.port.postMessage(merged, [merged.buffer]); // transfer ownership
      this._buffer = [];
      this._bufferSize = 0;
    }

    return true; // keep processor alive
  }
}

registerProcessor('pcm-capture-processor', PcmCaptureProcessor);
