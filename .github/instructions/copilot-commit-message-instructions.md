# Commit Guidelines (Simplified Version)

1. Basic Format:

   ```
   <type>(<scope>)?: <short description>
   ```

   - type: feat (new feature), fix (bug fix), docs (documentation), chore (maintenance), etc.
   - scope (optional): area affected, e.g., api, ui, build

2. After the description, add a blank line and optional body:

   - Explain reasons and implementation details.

3. Optional Footer:

   - BREAKING CHANGE: Describe major changes.
   - Or add `!` after the type to indicate breaking changes, e.g., `feat!: ...`.


<!-- 4. Additional Rule:

   - 請使用`繁體中文`撰寫敘述，專業術語請使用`英文`。

5. Examples:

   ```
   feat(api): 新增使用者登入功能

   增加 JWT 認證並回傳 token。

   fix: 修正註冊頁面輸入驗證問題

   chore!: 移除舊版 Node 支援

   BREAKING CHANGE: 不再支援 Node 8
   ``` -->

Keep messages concise and consistent for automation.
