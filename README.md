# AA Study Library — v2 replacement candidate

This branch is a **complete reader upgrade candidate**, built on the previous GitHub
reader and Study Lens. It is **not** the live ChatGPT-hosted Sites project.

## Included
- Navigation and search for Big Book, Steps, Traditions, Dictionary, Daily Reflections.
- Autoload sources from the owner's existing public GitHub repositories.
- Big Book printed-page navigation including actual print markers inside chapters.
- 60 independent study annotations, including 13 turning-point markers.
- Optional coral God/Higher Power highlights, green Daily Reflections excerpts
  and date popup, violet principles, and opt-in contextual pronouns.
- Markers for overlapping categories are combined without hiding one category.
- Previously working local Daily Reflections notes and font settings in the GitHub
  reader are untouched by these additive scripts.

## Pending before replacing the live ChatGPT Site
- Review mobile and desktop behavior in a real browser.
- Compare the earlier **57 approved Higher Power terms** against the current
  provisional vocabulary; the exact earlier checklist is not stored in this repo.
  Never automatically highlight `you` or `your`.
- Check all study annotation page markers against the printed books.
- Arrange migration of notes stored under the existing ChatGPT-site origin:
  separate website origins do not share browser localStorage.
- Deploy and switch the ChatGPT-hosted project using the Sites publishing workflow;
  GitHub connector cannot edit that distinct live project directly.

## Technical details
- `source-adapters.js`: converts full Big Book `pages[]` and 12&12 text into
  readable sections with their printed-page labels; later pages are chunked.
- `v2.js`: source routing, independent highlight layers, page navigator,
  and clickable Daily Reflections date tooltips.
- `study-map.json` and `study-lens.js`: interpretation notes, category filters,
  and click-to-print-page navigation.
- `v2.css`: page-number and highlighting presentation.
- `index.html`: loads all extensions after the original reader script.

## Verification performed
Static JavaScript compilation, valid JSON structure and study marker ranges,
source-adapter sample tests, and highlight-overlap/exclusion tests.
A remote interactive browser test timed out; **do not treat this as full UI QA**.

Independent study companion, not an official A.A. website.
