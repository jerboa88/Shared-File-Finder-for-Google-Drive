/*
 * Shared File Finder for Google Drive | @jerboa88 | MIT | (https://github.com/jerboa88/Shared-File-Finder-for-Google-Drive).
 */


const folderCache = {};


/*
 * Entry point
 */
function runSharedFileFinder() {
	const headerLabels = ['ID', 'Type', 'Path', 'Owners', 'Link'];
	const query = 'trashed = false and "me" in owners';
	const sheet = SpreadsheetApp.getActiveSheet();
	const numOfCols = headerLabels.length;
	const headerRow = sheet.getRange(1, 1, 1, numOfCols);
	let files;
	let pageToken = null;

	sheet.clear();
	sheet.setFrozenRows(1);
	headerRow.setValues([headerLabels]);
	headerRow.setFontWeight('bold');

	do {
		try {
			files = Drive.Files.list({
				q: query,
				maxResults: 1000,
				pageToken: pageToken,
			});

			if (!files.items || files.items.length === 0) {
				console.warn('No folders found.');

				return;
			}

			for (let i = 0; i < files.items.length; ++i) {
				const file = files.items[i];

				if (file.shared) {
					const fileType = file.mimeType === 'application/vnd.google-apps.folder' ? 'Folder' : 'File';
					const ownerEmailAddresses = file.owners.map(owner => owner.emailAddress);
					const filePath = getFilePath(file);

					sheet.appendRow([file.id, fileType, filePath, ownerEmailAddresses.toString(), file.alternateLink]);

					console.log(`${fileType} '${file.title}' is shared. Appended to sheet`);
				}
			}

			pageToken = files.nextPageToken;
		} catch (err) {
			console.error('Failed with error %s', err.message);
		}
	} while (pageToken);

	sheet.setColumnWidth(1, 50);
	sheet.autoResizeColumns(2, numOfCols);
	sheet.sort(3, true);
}


/*
 * Returns the full path of a file as a string.
 */
function getFilePath(file) {
	const folderNameList = [file.title];
	let parentId = getParentId(file);

	while (parentId) {
		let parentFolder;

		if (parentId in folderCache) {
			parentFolder = folderCache[parentId];
		} else {
			parentFolder = Drive.Files.get(parentId);
			folderCache[parentId] = parentFolder;
		}

		parentId = getParentId(parentFolder);

		folderNameList.push(parentFolder.title);
	}

	return `/${folderNameList.reverse().join('/')}`;
}


/*
 * Returns the ID of the parent folder of a file.
 */
function getParentId(file) {
	return file.parents.length < 1 || file.parents[0].isRoot ? null : file.parents[0].id;
}
