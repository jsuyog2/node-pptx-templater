# New Documentation Site — Design and Technical Requirements

## Goal
Create a polished documentation website from scratch for `node-pptx-templater`, replacing the old GitHub Pages docs. The site must be maintainable, responsive, accessible, and based on the real library API.

## Inspect before choosing the stack
- Inspect `package.json`, existing frontend assets, repository build conventions, and GitHub Pages workflow.
- Use Tailwind CSS and GSAP animations as requested.
- Prefer a lightweight site setup compatible with the existing repository and deployment target. Do not add a large framework without a clear benefit.
- If no site framework exists, choose a simple maintainable static site setup that supports Tailwind compilation, GSAP, routing or section navigation, and GitHub Pages base paths.
- Keep docs source in a clearly named new location, such as `website/` or `site/`, and configure output/deployment consistently. Do not recreate the removed old `docs/` directory as a build output tracked by source control unless deployment explicitly requires it.

## Visual design
- Glassmorphism: translucent panels, restrained backdrop blur, subtle borders, layered gradients, and soft shadows.
- Premium developer-tool aesthetic. Prioritize readability, high contrast, spacing, and performance over excessive effects.
- Add a working Light/Dark theme toggle. Persist preference only if the chosen implementation supports it safely; otherwise use a reliable in-page toggle and system preference as the initial default.
- Use GSAP for purposeful, subtle entrance, section reveal, and interaction animations. Respect `prefers-reduced-motion`; disable or reduce motion when requested by the OS.
- Responsive layouts for mobile, tablet, and desktop.
- Include accessible semantic landmarks, keyboard-visible focus, sufficient contrast in both themes, accessible names for icon-only controls, and useful alt text.
- Ensure content remains readable if blur or animation is unsupported.
- Avoid scroll hijacking, perpetual distracting motion, giant animation bundles, and layout shifts.

## Required content and navigation
- Hero section explaining what the library does with primary install command.
- Quick start with verified code.
- Navigation for guides, API/reference, examples, input/output formats, slides, text, shapes, images, tables, charts, hyperlinks, serialization, troubleshooting, and feature status where implemented/documented.
- Feature catalog grouped by subsystem and accurately labeled as supported, partial, planned, or unverified based on audit evidence.
- A single-file JSON-driven showcase example and exact command to run it.
- Clear installation instructions and Node.js compatibility requirements verified from `package.json`.
- Links to npm and GitHub:
  - npm: https://www.npmjs.com/package/node-pptx-templater
  - GitHub: https://github.com/jsuyog2/node-pptx-templater
- Provide source-code links to the actual repository paths when possible.
- Include copy-to-clipboard buttons for install commands and key snippets; ensure controls work and provide accessible feedback.
- Add search/filter for docs only if it can be implemented robustly without disproportionate complexity.
- Include footer with npm, GitHub, license, and version sourced from project metadata.

## Documentation accuracy
- Verify every method name, signature, option, return value, and example against source exports and tests.
- Do not publish the entire 86-item list as implemented unless the audit confirms each item.
- Explain known limitations and unsupported formats honestly.
- Never invent performance metrics, compatibility guarantees, downloads, badges, release versions, or coverage percentages.

## Deployment
- Configure the new site to deploy through GitHub Pages using the repository's actual deployment model.
- Support the correct repository base path (`/node-pptx-templater/`) if deploying as a project site.
- Use repository-relative asset paths or the framework's base URL mechanism so CSS, JS, images, and links work on GitHub Pages.
- Update workflow permissions, artifact paths, and build commands consistently.
- Test the built site locally and verify that internal navigation, theme toggle, animations, and assets resolve under the configured base path.
- Do not claim deployment is live unless a deployment was actually run and verified.
