var fs = require("fs");
var Q = String.fromCharCode(39);
var BT = String.fromCharCode(96);
var D = String.fromCharCode(36);
var content = "" +
"// Request Service
" +
"// Asset request workflow management
" +
"
" +
"import { SPFI } from " + Q + "@pnp/sp" + Q + ";
" +
"import " + Q + "@pnp/sp/webs" + Q + ";
" +
"import " + Q + "@pnp/sp/lists" + Q + ";
" +
"import " + Q + "@pnp/sp/items" + Q + ";
" +
"import " + Q + "@pnp/sp/items/get-all" + Q + ";
" +
"import " + Q + "@pnp/sp/site-users/web" + Q + ";
" +
"import { IAssetRequest, AssetCategory } from " + Q + "../models/IAsset" + Q + ";
" +
"import { AM_LISTS } from " + Q + "../constants/SharePointListNames" + Q + ";
" +
"
";
console.log("helper works, length: " + content.length);
process.exit(0);