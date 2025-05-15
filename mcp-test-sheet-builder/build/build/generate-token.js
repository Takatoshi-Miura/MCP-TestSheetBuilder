import { generateToken } from "./auth.js";
// トークン生成コマンドを実行
console.log("Google API認証トークン生成ツール");
console.log("================================");
console.log("認証プロセスを開始します...");
generateToken()
    .then(() => {
    console.log("処理が完了しました。");
    process.exit(0);
})
    .catch(error => {
    console.error("エラーが発生しました:", error);
    process.exit(1);
});
