# Supriyo’s Search Lab

SEO, Organic Marketing & Automation Experiments, built with [Astro](https://astro.build) and published free on GitHub Pages.

## Write a new post

1. Create a file in `src/content/blog/`. The file name becomes the URL, so use the keyword:
   `src/content/blog/python-bulk-rename-files.md` → `/blog/python-bulk-rename-files/`
2. Start it with:

   ```md
   ---
   title: 'How to Bulk Rename Files with Python'
   description: '120–155 characters that make someone want to click in Google.'
   pubDate: '2026-10-22'
   tags: ['Python', 'File Automation']
   draft: true
   ---
   ```

3. Write in Markdown below it. Remove `draft: true` when it's ready to go live.
4. When you edit an old post, add `updatedDate: '2026-11-05'`. Google likes fresh content.

**Never rename a published post's file.** That changes its URL and loses its rankings.

## Commands

| Command           | What it does                                   |
| ----------------- | ---------------------------------------------- |
| `npm run dev`     | Preview at http://localhost:4321 (shows drafts) |
| `npm run build`   | Build the live site into `dist/`               |

## Publish

```
git add .
git commit -m "New post: <title>"
git push
```

GitHub builds and publishes the site in about a minute.

## When you buy your domain

1. In `astro.config.mjs`, change `site` to your domain.
2. In `public/robots.txt`, change the sitemap URL.
3. Add the domain under the repository's **Settings → Pages → Custom domain**.

GitHub then redirects the old `github.io` address to your domain automatically.
