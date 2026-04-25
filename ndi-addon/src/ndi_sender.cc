/*
 * SermonCast NDI Sender — Native Node.js Addon
 * Uses the official NDI SDK v6 to broadcast the projection window as a native NDI source.
 * Visible in OBS, vMix, Wirecast, and any NDI-capable software on the same network.
 *
 * API (synchronous, main-process use):
 *   createSender(name, width, height, fpsN, fpsD) → true | throws
 *   sendBGRA(buffer)                               → true | throws
 *   destroySender()                                → undefined
 *   isReady()                                      → boolean
 */

#include <napi.h>
#include <cstring>
#include <mutex>
#include <string>

#include "Processing.NDI.Lib.h"

namespace {
    std::mutex g_mutex;
    NDIlib_send_instance_t g_sender = nullptr;

    int g_width  = 1920;
    int g_height = 1080;
    int g_fpsN   = 30000;
    int g_fpsD   = 1000;

    bool g_initialized = false;
    bool g_init_ok     = false;

    bool ensure_ndi_initialized() {
        if (!g_initialized) {
            g_init_ok     = NDIlib_initialize();
            g_initialized = true;
        }
        return g_init_ok;
    }
}

// createSender(name: string, width: int, height: int, fpsN: int, fpsD: int) → boolean
Napi::Value CreateSender(const Napi::CallbackInfo& info) {
    Napi::Env env = info.Env();

    if (info.Length() < 5) {
        Napi::TypeError::New(env, "createSender(name, width, height, fpsN, fpsD)").ThrowAsJavaScriptException();
        return env.Null();
    }

    std::string name = info[0].As<Napi::String>().Utf8Value();
    g_width  = info[1].As<Napi::Number>().Int32Value();
    g_height = info[2].As<Napi::Number>().Int32Value();
    g_fpsN   = info[3].As<Napi::Number>().Int32Value();
    g_fpsD   = info[4].As<Napi::Number>().Int32Value();

    if (!ensure_ndi_initialized()) {
        Napi::Error::New(env, "NDIlib_initialize() failed — is NDI Runtime installed?").ThrowAsJavaScriptException();
        return env.Null();
    }

    std::lock_guard<std::mutex> lock(g_mutex);

    // Destroy existing sender before creating new one
    if (g_sender) {
        NDIlib_send_destroy(g_sender);
        g_sender = nullptr;
    }

    NDIlib_send_create_t desc;
    std::memset(&desc, 0, sizeof(desc));
    desc.p_ndi_name  = name.c_str();
    desc.p_groups    = nullptr;
    desc.clock_video = true;   // clock-accurate video timing
    desc.clock_audio = false;

    g_sender = NDIlib_send_create(&desc);
    if (!g_sender) {
        Napi::Error::New(env, "NDIlib_send_create() failed").ThrowAsJavaScriptException();
        return env.Null();
    }

    return Napi::Boolean::New(env, true);
}

// sendBGRA(buffer: Buffer) → boolean
// buffer must be exactly width * height * 4 bytes (BGRA format)
Napi::Value SendBGRA(const Napi::CallbackInfo& info) {
    Napi::Env env = info.Env();

    if (info.Length() < 1 || !info[0].IsBuffer()) {
        Napi::TypeError::New(env, "sendBGRA(buffer: Buffer)").ThrowAsJavaScriptException();
        return env.Null();
    }

    std::lock_guard<std::mutex> lock(g_mutex);

    if (!g_sender) {
        Napi::Error::New(env, "NDI sender not created — call createSender() first").ThrowAsJavaScriptException();
        return env.Null();
    }

    Napi::Buffer<uint8_t> buf = info[0].As<Napi::Buffer<uint8_t>>();
    const size_t expected = (size_t)g_width * (size_t)g_height * 4;

    if (buf.Length() < expected) {
        Napi::Error::New(env, "Buffer too small for frame size").ThrowAsJavaScriptException();
        return env.Null();
    }

    NDIlib_video_frame_v2_t frame;
    std::memset(&frame, 0, sizeof(frame));

    frame.xres                 = g_width;
    frame.yres                 = g_height;
    frame.FourCC               = NDIlib_FourCC_type_BGRA;
    frame.frame_rate_N         = g_fpsN;
    frame.frame_rate_D         = g_fpsD;
    frame.picture_aspect_ratio = (float)g_width / (float)g_height;
    frame.frame_format_type    = NDIlib_frame_format_type_progressive;
    frame.timecode             = NDIlib_send_timecode_synthesize;
    frame.line_stride_in_bytes = g_width * 4;
    frame.p_data               = buf.Data();

    NDIlib_send_send_video_v2(g_sender, &frame);

    return Napi::Boolean::New(env, true);
}

// destroySender() → undefined
Napi::Value DestroySender(const Napi::CallbackInfo& info) {
    Napi::Env env = info.Env();
    std::lock_guard<std::mutex> lock(g_mutex);
    if (g_sender) {
        NDIlib_send_destroy(g_sender);
        g_sender = nullptr;
    }
    NDIlib_destroy();
    g_initialized = false;
    return env.Undefined();
}

// isReady() → boolean — true if sender is created and ready
Napi::Value IsReady(const Napi::CallbackInfo& info) {
    return Napi::Boolean::New(info.Env(), g_sender != nullptr);
}

Napi::Object Init(Napi::Env env, Napi::Object exports) {
    exports.Set("createSender",  Napi::Function::New(env, CreateSender));
    exports.Set("sendBGRA",      Napi::Function::New(env, SendBGRA));
    exports.Set("destroySender", Napi::Function::New(env, DestroySender));
    exports.Set("isReady",       Napi::Function::New(env, IsReady));
    return exports;
}

NODE_API_MODULE(ndi_sender, Init)
