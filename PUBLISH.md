# Publishing to GitHub Pages (instructions for Claude Code)

Publish the contents of this folder as a static GitHub Pages site. There is no build step.

1. From this folder, create a git repository and commit everything, including the empty `.nojekyll` file:
   git init -b main && git add . && git commit -m "Federer 54: standalone web version"
2. Create a public repository named `federer54` on the user's GitHub account and push:
   gh repo create federer54 --public --source=. --remote=origin --push
3. Turn on GitHub Pages from the `main` branch, root folder:
   gh api -X POST repos/{owner}/federer54/pages -f "source[branch]=main" -f "source[path]=/"
4. The site will be at https://<owner>.github.io/federer54/ after a minute or two. Check the status with:
   gh api repos/{owner}/federer54/pages

Notes
- Photos are intentionally not included. Do not add images to this repository until the owner has confirmed the rights to publish them.
- `data.js` is about 19 MB; this is within GitHub's limits (GitHub warns above 50 MB).
- If the repository already exists, skip step 2 and push to it instead: git remote add origin <url> && git push -u origin main
