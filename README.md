<!-- Project Header -->
<div align="center">
  <!-- <img class="projectLogo" src="https://via.placeholder.com/256.jpg" alt="Project logo" title="Project logo" width="256"> -->

  <h1 class="projectName">Shared File Finder for Google Drive</h1>

  <p class="projectBadges">
    <img src="https://img.shields.io/badge/type-Apps_Script-2196f3.svg" alt="Project type" title="Project type">
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
 3. Copy and paste the code from this file into the Script Editor.
 4. Save the project and run the `runSharedFileFinder` function.

**Notes:**
  - Files must be owned by the current Google Drive user.
  - If a folder is shared, both the folder and its files may be shown in the list.
  - There may be bugs. Use at your own risk.


## Contributing
Contributions, issues, and forks are welcome but this is a hobby project so don't expect too much from it. [SemVer](http://semver.org/) is used for versioning.


## License
Inspired by a similar script by @danjargold (https://gist.github.com/danjargold/c6542e68fe3a3b46eeb0172f914641bc) and @woodwardtw (https://gist.github.com/woodwardtw/22a199ecca73ff15a0eb). This version uses the Drive API v2 to get info for multiple files at once (which makes it substantially faster).

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.
