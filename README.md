<!-- Project Header -->
<div align="center">
  <img class="projectLogo" src="screenshot.png" alt="Project logo" title="Project logo" width="256">

  <h1 class="projectName">Shared File Finder for Google Drive</h1>

  <p class="projectBadges">
    <img src="https://johng.io/badges/category/Script.svg" alt="Project category" title="Project category">
    <img src="https://img.shields.io/github/languages/top/jerboa88/Shared-File-Finder-for-Google-Drive.svg" alt="Language" title="Language">
    <img src="https://img.shields.io/github/repo-size/jerboa88/Shared-File-Finder-for-Google-Drive.svg" alt="Repository size" title="Repository size">
    <a href="LICENSE">
      <img src="https://img.shields.io/github/license/jerboa88/Shared-File-Finder-for-Google-Drive.svg" alt="Project license" title="Project license"/>
    </a>
  </p>

  <p class="projectDesc">
    An Apps Script that finds all files/folders on Google Drive that are shared with others and adds them to a Google Sheet.
  </p>

  <br/>
</div>


## Usage
 1. Create a new Google Sheet.
 2. Open the Script Editor (Extensions > Apps Script).
 3. Copy and paste the code from [shared-file-finder.js](shared-file-finder.js) into the Script Editor.
 4. Enable `Drive API` in the `Advanced Services` list for the project (see [here](https://developers.google.com/apps-script/guides/services/advanced#enable_advanced_services) for instructions). The API version should be v2 and the identifier should be `Drive`.
 5. Save the project and run the `runSharedFileFinder()` function.

**Notes:**
  - Files must be owned by the current Google Drive user.
  - If a folder is shared, both the folder and its files will be shown in the list.
  - There may be bugs. Use at your own risk.


## Contributing
Contributions, issues, and forks are welcome. [SemVer](http://semver.org/) is used for versioning.


## License
Inspired by a similar script by @danjargold (https://gist.github.com/danjargold/c6542e68fe3a3b46eeb0172f914641bc) and @woodwardtw (https://gist.github.com/woodwardtw/22a199ecca73ff15a0eb). This version uses the Drive API v2 to get info for multiple files at once (which makes it substantially faster).

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.


## 💕 Funding

Find this project useful? [Sponsoring me](https://johng.io/funding) will help me cover costs and **_commit_** more time to open-source.

If you can't donate but still want to contribute, don't worry. There are many other ways to help out, like:

- 📢 reporting (submitting feature requests & bug reports)
- 👨‍💻 coding (implementing features & fixing bugs)
- 📝 writing (documenting & translating)
- 💬 spreading the word
- ⭐ starring the project

I appreciate the support!
