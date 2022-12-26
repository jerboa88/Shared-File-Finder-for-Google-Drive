/*
 * Shared File Finder for Google Drive | @jerboa88 | MIT | (https://github.com/jerboa88/Shared-File-Finder-for-Google-Drive).
 */


/*
 * Class for caching reusuable values.
 */
class Cache {
	constructor() {
		this.cache = {};
	}

	get(key, callback) {
		if (key in this.cache) {
			return this.cache[key];
		}

		const value = callback();

		this.cache[key] = value;

		return value;
	}
}


/*
 * Entry point
 */
function runSharedFileFinder() {
	// Constants
	const resultsSheetName = 'Shared Files';
	const folderBgColor = '#FFF8E1';
	const fileBgColor = '#E0F7FA';
	const headerLabels = ['ID', 'Type', 'Path', 'Owners'];
	const query = 'trashed = false and "me" in owners';
	const chunkSize = 1000;
	const isDebugMode = false;

	// Runtime
	const folderCache = new Cache();
	const cellIconCache = new Cache();
	const fileSummaryList = [];
	let files;
	let pageToken = null;
	let numOfFilesProcessed = 0;

	do {
		try {
			files = Drive.Files.list({
				q: query,
				maxResults: chunkSize,
				pageToken: pageToken,
			});

			if (!files.items || files.items.length === 0) {
				console.warn('No folders found.');

				return;
			}

			for (let i = 0; i < files.items.length; ++i) {
				const file = files.items[i];

				if (file.shared) {
					const isFolder = file.mimeType === 'application/vnd.google-apps.folder';
					const ownerEmailAddresses = file.owners.map(owner => owner.emailAddress);
					const filePath = getFilePath(folderCache, file);

					fileSummaryList.push({
						id: file.id,
						isFolder: isFolder,
						iconLink: file.iconLink,
						path: filePath,
						ownerEmailAddresses: ownerEmailAddresses.toString(),
						link: file.alternateLink,
					});

					console.log(`${isFolder ? 'Folder' : 'File'} '${file.title}' is shared. Adding to list`);
				}
			}

			numOfFilesProcessed += files.items.length;

			console.log(`${numOfFilesProcessed} files processed`);

			pageToken = files.nextPageToken;
		} catch (err) {
			console.error(`Failed with error: ${err.message}`);
		}
	} while (isDebugMode ? false : pageToken);

	const numOfRows = fileSummaryList.length;
	const numOfCols = headerLabels.length;
	const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
	const resultsSheet = createResultsSheet(activeSpreadsheet, resultsSheetName, numOfRows, numOfCols, headerLabels);
	const dataRange = resultsSheet.getRange(2, 1, numOfRows, numOfCols);
	const fileTypeDataRange = dataRange.offset(0, 1, fileSummaryList.length, 1);

	populateResultsSheet(dataRange, fileTypeDataRange, cellIconCache, fileSummaryList, folderBgColor, fileBgColor);
	formatResultsSheet(activeSpreadsheet, resultsSheet, fileTypeDataRange, numOfCols);
}


/*
 * Returns the full path of a file as a string.
 */
function getFilePath(folderCache, file) {
	const folderNameList = [file.title];
	let parentId = getParentId(file);

	while (parentId) {
		const parentFolder = folderCache.get(parentId, () => Drive.Files.get(parentId));

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
 * Returns a CellImage object with the given image URL.
 */
function createCellImage(imageUrl) {
	return SpreadsheetApp.newCellImage()
		.setSourceUrl(imageUrl)
		.setAltTextTitle(imageUrl)
		.setAltTextDescription('File type icon')
		.build();
}


/*
 * Creates a new sheet to display the progress of the script.
 */
function createResultsSheet(activeSpreadsheet, sheetName, numOfRows, numOfCols, headerLabels) {
	sheetName = `${sheetName} (${getDateString()})`;

	let resultsSheet = activeSpreadsheet.getSheetByName(sheetName);

	if (resultsSheet != null) {
		activeSpreadsheet.deleteSheet(resultsSheet);
	}

	resultsSheet = activeSpreadsheet.insertSheet(sheetName);

	resizeSheet(resultsSheet, numOfRows, numOfCols);

	resultsSheet.setFrozenRows(1);
	resultsSheet.setColumnWidth(1, 50);

	const headerRowRange = resultsSheet.getRange(1, 1, 1, numOfCols);

	headerRowRange.setValues([headerLabels]);
	headerRowRange.setFontWeight('bold');

	return resultsSheet;
}


/*
 * Resize a sheet to the given number of rows and columns.
 */
function resizeSheet(sheet, numOfRows, numOfCols) {
	if (sheet.getMaxRows() < numOfRows) {
		sheet.insertRowsAfter(sheet.getMaxRows(), numOfRows - sheet.getMaxRows());
	} else if (sheet.getMaxRows() > numOfRows) {
		sheet.deleteRows(numOfRows + 1, sheet.getMaxRows() - numOfRows);
	}

	if (sheet.getMaxColumns() < numOfCols) {
		sheet.insertColumnsAfter(sheet.getMaxColumns(), numOfCols - sheet.getMaxColumns());
	} else if (sheet.getMaxColumns() > numOfCols) {
		sheet.deleteColumns(numOfCols + 1, sheet.getMaxColumns() - numOfCols);
	}
}


/*
 * Map the file summary list to the results sheet.
 */
function populateResultsSheet(dataRange, fileTypeDataRange, cellIconCache, fileSummaryList, folderBgColor, fileBgColor) {
	const sortedFileSummaryList = fileSummaryList.sort(({ path: path1 }, { path: path2 }) => path1.localeCompare(path2));
	const blankRichTextValue = createRichTextValue('');

	dataRange.setRichTextValues(
		sortedFileSummaryList.map(fileSummary => [
			createRichTextValue(fileSummary.id),
			blankRichTextValue,
			createRichTextValue(fileSummary.path, fileSummary.link),
			createRichTextValue(fileSummary.ownerEmailAddresses)
		]));

	sortedFileSummaryList.forEach((fileSummary, i) => {
		dataRange.offset(i, 0, 1).setBackground(fileSummary.isFolder ? folderBgColor : fileBgColor);
	});

	const iconLinks = sortedFileSummaryList.map(({ iconLink }) => {
		return cellIconCache.get(iconLink, () => [createCellImage(iconLink)]);
	});

	console.log('Loading file type icons. This may take a minute to complete');

	fileTypeDataRange.setValues(iconLinks);
}


/*
 * Format the results sheet after data has been added.
 */
function formatResultsSheet(activeSpreadsheet, resultsSheet, fileTypeDataRange, numOfCols) {
	resultsSheet.autoResizeColumns(2, numOfCols - 1);
	fileTypeDataRange.setHorizontalAlignment('center');
	activeSpreadsheet.setActiveSheet(resultsSheet);
}


/*
 * Returns the current date as a string.
 */
function getDateString() {
	const date = new Date();

	return `${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`;
}
