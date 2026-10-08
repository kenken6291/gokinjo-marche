# ご近所マルシェ＆助け合い広場（gokinjo-marche）

GitHub Pages（index.html + config.js） + GAS（Code.gs） + スプレッドシート + Google Drive + Gemini API

## ① スプレッドシートのテーブル定義

`setup()` を実行すると、下記4シートが1行目の見出し付きで自動作成されます。

### Users
| 列 | 内容 |
|---|---|
| userId | `U_yyMMddHHmmss_xxxxxx` |
| email | 小文字で保存（ログインID） |
| passwordHash | SHA-256（salt＋pepper、500回ストレッチ）の16進文字列 |
| salt | ユーザーごとのUUID |
| isTemp | 仮パスワード中なら TRUE（本パスワード設定まで他機能ロック） |
| area | 地域（〇〇町など） |
| nickname | 表示名 |
| createdAt | 登録日時 |
| failCount | ログイン連続失敗回数（5回で15分ロック） |
| lockedUntil | ロック解除時刻 |

### Posts（おすそ分け・物々交換・助け合い）
| 列 | 内容 |
|---|---|
| postId / userId | 投稿ID / 投稿者 |
| type | `exchange`（物々交換系） / `help`（助け合い） |
| category | `vegetable` 🥬 / `lend` 🪑 / `barter` 🎁 / `help` 🛠️ |
| title / detail | タイトル / 詳細 |
| imageUrl | Drive画像のサムネイルURL |
| status | `open` 募集中 / `talking` 相談中 / `done` 受け渡し済み |
| area | 地域 |
| contacts | 連絡先JSON `{zoom, meet, line, facebook, email, phone}`（会員のみに返却） |
| createdAt | 投稿日時 |

### Events（広場のイベント）
| 列 | 内容 |
|---|---|
| eventId / organizerId | イベントID / 主催者 |
| category | `marche` 🎪 / `flea` 📦 / `offkai` ☕ / `matsuri` 🏮 / `help` 🤝 |
| title / datetime / place / detail | タイトル / 日時(ISO) / 開催場所 / 案内文 |
| imageUrl | 写真・チラシ画像 |
| joinCount / likeCount | 🙋 参加表明数 / 👍 楽しそう数 |
| slots | 持ち寄り・出店枠JSON `[{id, label, capacity, members:[userId]}]` |
| contacts | 外部通話・連絡先JSON（Postsと同じ形） |
| area / createdAt | 地域 / 作成日時 |

### Matches（参加・リアクション・申込）
| 列 | 内容 |
|---|---|
| matchId | `M_...` |
| targetId | eventId または postId |
| targetType | `event` / `post` |
| applicantId | 申込者のuserId |
| kind | `join` 参加表明 / `like` 楽しそう / `request` おすそ分け申込 |
| note | 申込時のひとこと |
| createdAt | 成立日時 |

> セッションは CacheService（6時間、利用のたびに延長）で管理するためシートは不要です。

## セットアップ

1. スプレッドシートを新規作成 → 拡張機能 › Apps Script に `Code.gs` を貼り付け
2. プロジェクトの設定（⚙）
   - タイムゾーン：**(GMT+09:00) 東京**
   - スクリプトプロパティ：`GEMINI_API_KEY`（必須）、`GEMINI_MODEL`（任意。未設定なら `gemini-3.8-flash`）
3. `setup()` を実行して権限を承認（シート・Driveフォルダ・PEPPERが作られます）
4. デプロイ › 新しいデプロイ › ウェブアプリ（実行：自分／アクセス：全員）
5. 発行URLを `config.js` の `GAS_URL` に貼り付け → `index.html` と `config.js` をリポジトリ `gokinjo-marche` に push → Pages を有効化

**Code.gs を直したら、「デプロイを管理」› 編集 › バージョン「新バージョン」で再デプロイが必要です**（URLはそのまま）。

## API一覧（POST・`Content-Type: text/plain`）

| action | ログイン | 内容 |
|---|---|---|
| register / login / resetPassword | 不要 | 登録（仮PWメール） / ログイン / 仮PW再発行 |
| listEvents / listPosts | 不要 | 一覧（連絡先はログイン時のみ） |
| me / logout / changePassword / updateProfile | 必要 | 会員情報 |
| uploadImage | 必要 | Base64画像 → Drive保存（公開リンク） |
| createEvent / updateEvent / deleteEvent | 必要 | イベント（編集・削除は主催者のみ） |
| toggleJoin / toggleLike / toggleSlot | 必要 | 参加表明 / 楽しそう / 出店枠（主催者へメール通知） |
| eventMembers | 必要 | 参加者一覧（主催者・参加者のみ） |
| createPost / updatePost / deletePost / requestPost | 必要 | おすそ分け（申込は投稿者へメール通知） |
| aiWrite | 必要 | Geminiで見出し・タイトル・告知文を生成 |

## メモ

- Google Meet はAPIで会議を自動発行できない（OAuthが必要）ため、「新しく作る」で `meet.google.com/new` を開き、表示されたリンクを貼り付ける方式です。
- Gmail の送信上限は無料アカウントで1日100通前後です。
- 画像はブラウザ側で長辺1280pxに縮小してから送ります（5MBまで）。
