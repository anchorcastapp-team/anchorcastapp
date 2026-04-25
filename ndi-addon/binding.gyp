{
  "targets": [
    {
      "target_name": "ndi_sender",
      "sources": [ "src/ndi_sender.cc" ],
      "include_dirs": [
        "<!@(node -p \"require('node-addon-api').include\")",
        "C:/Program Files/NDI/NDI 6 SDK/Include"
      ],
      "libraries": [
        "C:/Program Files/NDI/NDI 6 SDK/Lib/x64/Processing.NDI.Lib.x64.lib"
      ],
      "defines": [ "NAPI_CPP_EXCEPTIONS" ],
      "msvs_settings": {
        "VCCLCompilerTool": {
          "ExceptionHandling": 1,
          "AdditionalOptions": [ "/std:c++17" ]
        }
      },
      "conditions": [
        ["OS=='win'", {
          "copies": [
            {
              "destination": "<(PRODUCT_DIR)",
              "files": [
                "C:/Program Files/NDI/NDI 6 SDK/Bin/x64/Processing.NDI.Lib.x64.dll"
              ]
            }
          ]
        }]
      ]
    }
  ]
}
