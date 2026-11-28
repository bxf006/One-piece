# Third-party notices

This project includes JSZip 3.10.1 (https://stuk.github.io/jszip/), used under the MIT license. Its license is reproduced in `vendor/JSZIP-LICENSE.md` and its copyright notice is preserved in `vendor/jszip.min.js`. JSZip includes pako compression code under the MIT license; the bundled license file includes the relevant notices.

This project includes SheetJS Community Edition 0.20.3 (https://docs.sheetjs.com/), used under the Apache License, Version 2.0. Its license is reproduced in `vendor/SHEETJS-LICENSE.txt`; its copyright notice is preserved in `vendor/xlsx.full.min.js`. The bundle was copied from the user-provided `xlsx.full.min.js` and its version was verified. The authoritative distribution is https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js.

The chart renderer, data adapters, application code and styles are authored for this project. Third-party runtime dependencies are separate files in `vendor/`, referenced by relative paths from `index.html`. Application code and styles remain embedded in the HTML. No runtime CDN or remote API is used.
