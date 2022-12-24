/*
 * Shared File Finder for Google Drive | @jerboa88 | MIT | (https://github.com/jerboa88/Shared-File-Finder-for-Google-Drive).
 */


const folderCache = {};
const fileSummaryList = [];


/*
 * Entry point
 */
function runSharedFileFinder() {
	// Config vars
	const folderLabel = 'Folder';
	const fileLabel = 'File';
	const headerLabels = ['ID', 'Type', 'Path', 'Owners'];
	const query = 'trashed = false and "me" in owners';
	const isDebugMode = false;

	// Runtime vars
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
					const fileType = file.mimeType === 'application/vnd.google-apps.folder' ? folderLabel : fileLabel;
					const ownerEmailAddresses = file.owners.map(owner => owner.emailAddress);
					const filePath = getFilePath(file);

					fileSummaryList.push({
						id: file.id,
						type: fileType,
						path: filePath,
						ownerEmailAddresses: ownerEmailAddresses.toString(),
						link: file.alternateLink,
					});

					console.log(`${fileType} '${file.title}' is shared. Adding to list`);
				}
			}

			pageToken = files.nextPageToken;
		} catch (err) {
			console.error(`Failed with error: ${err.message}`);
		}
	} while (isDebugMode ? false : pageToken);

	const dataRange = sheet.getRange(2, 1, fileSummaryList.length, numOfCols);
	const folderRule = createConditionalFormatRule(dataRange, folderLabel, '#FFF8E1');
	const fileRule = createConditionalFormatRule(dataRange, fileLabel, '#E0F7FA');

	dataRange.setRichTextValues(
		fileSummaryList
			.sort(({ path: path1 }, { path: path2 }) => path1.localeCompare(path2))
			.map(fileSummary => [
				createRichTextValue(fileSummary.id),
				createRichTextValue(fileSummary.type),
				createRichTextValue(fileSummary.path, fileSummary.link),
				createRichTextValue(fileSummary.ownerEmailAddresses)
			]));

	sheet.setColumnWidth(1, 50);
	sheet.autoResizeColumns(2, numOfCols);
	sheet.setConditionalFormatRules([folderRule, fileRule]);
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


/*
 * Returns a RichTextValue object with the given text string and (optional) link URL.
 */
function createRichTextValue(text, linkUrl = null) {
	const rtv = SpreadsheetApp.newRichTextValue().setText(text);

	if (linkUrl) {
		rtv.setLinkUrl(linkUrl);
	}

	return rtv.build();
}


/*
 * Returns a ConditionalFormatRule object to highlight rows in different colors.
 */
function createConditionalFormatRule(range, strToMatch, bgColor) {
	return SpreadsheetApp.newConditionalFormatRule()
		.whenFormulaSatisfied(`=$B2 = "${strToMatch}"`)
		.setBackground(bgColor)
		.setRanges([range])
		.build();
}
