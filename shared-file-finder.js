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

	get(key, callback = () => { }) {
		if (key in this.cache) {
			return this.cache[key];
		}

		const value = callback();

		this.cache[key] = value;

		return value;
	}

	set(key, value) {
		this.cache[key] = value;
	}
}


/*
 * Entry point
 */
function runSharedFileFinder() {
	// Constants
	const resultsSheetName = 'Shared Files';
	const colorMap = {
		folderBg: '#FFF8E1',
		fileBg: '#E0F7FA',
		readerFont: '#000000',
		writerFont: '#B71C1C',
		commenterFont: '#1B5E20',
	};
	const headerLabels = ['ID', '', 'Path', 'Users'];
	const query = 'trashed = false and "me" in owners';
	const chunkSize = 1000;
	const isDebugMode = false;

	// Runtime
	const totalNumOfBytesUsed = parseInt(Drive.About.get().quotaBytesUsed);
	const folderCache = new Cache();
	const folderPathCache = new Cache();
	const cellIconCache = new Cache();
	const fileSummaryList = [];
	let files;
	let pageToken = null;
	let numOfFilesProcessed = 0;
	let numOfBytesProcessed = 0;

	console.log(`${totalNumOfBytesUsed} bytes of storage used`);
	console.log('Processing files...');

	do {
		try {
			files = Drive.Files.list({
				q: query,
				maxResults: chunkSize,
				spaces: 'drive',
				fields: 'items(id, mimeType, iconLink, alternateLink, title, shared, parents(isRoot, id), quotaBytesUsed), nextPageToken',
				pageToken: pageToken
			});

			if (!files.items || files.items.length === 0) {
				console.warn('No files or folders found.');

				return;
			}

			for (let i = 0; i < files.items.length; ++i) {
				const file = files.items[i];

				if (file.shared) {
					const isFolder = file.mimeType === 'application/vnd.google-apps.folder';

					fileSummaryList.push({
						id: file.id,
						isFolder: isFolder,
						iconLink: file.iconLink,
						path: getFilePath(folderCache, folderPathCache, file),
						users: getUserList(file.id),
						link: file.alternateLink,
					});

					console.log(`${isFolder ? 'Folder' : 'File'} '${file.title}' is shared. Adding to list`);
				}

				numOfBytesProcessed += parseInt(file.quotaBytesUsed);
			}

			numOfFilesProcessed += files.items.length;

			console.log(`${numOfFilesProcessed} files processed (~${(numOfBytesProcessed * 100 / totalNumOfBytesUsed).toFixed(2)}%)`);

			pageToken = files.nextPageToken;
		} catch (err) {
			console.error(`Failed with error: ${err.message}`);
		}
	} while (isDebugMode ? false : pageToken);

	console.log(`Done processing files`);

	createResultsSheet(resultsSheetName, headerLabels, fileSummaryList, cellIconCache, colorMap);
}


/*
 * Returns a list of users with access to a file.
 */
function getUserList(fileId) {
	let permissionsList = Drive.Permissions.list(fileId, {
		fields: 'items(emailAddress, role)'
	});

	if (!permissionsList.items || permissionsList.items.length === 0) {
		console.warn(`No permissions found for file ${fileId}`);

		return;
	}

	let userList = [];

	for (let i = 0; i < permissionsList.items.length; ++i) {
		const { emailAddress, role } = permissionsList.items[i];

		if (role === 'reader' || role === 'writer' || role === 'commenter') {
			userList.push({
				emailAddress: emailAddress ? emailAddress : 'Anyone with the link',
				role: role
			});
		}
	}

	return userList;
}


/*
 * Returns the full path of a file as a string.
 */
function getFilePath(folderCache, folderPathCache, file) {
	const parentFoldersList = [{
		id: file.id,
		title: file.title
	}];
	let parentId = getParentId(file);
	const parentPath = folderPathCache.get(parentId);

	if (parentPath) {
		return `${parentPath}/${file.title}}`;
	}

	while (parentId) {
		const parentFolder = folderCache.get(parentId, () => Drive.Files.get(parentId, {
			fields: 'parents(isRoot, id), title'
		}));

		parentFoldersList.push({
			id: parentId,
			title: parentFolder.title
		});

		parentId = getParentId(parentFolder);
	}

	let currentFolderPath = '';

	for (let i = parentFoldersList.length - 1; i >= 0; --i) {
		currentFolderPath += `/${parentFoldersList[i].title}`;

		folderPathCache.set(parentFoldersList[i].id, currentFolderPath);
	}

	return `${currentFolderPath}/${file.title}`;
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
 * Returns a RichTextValue object built from a list of users and their roles.
 */
function createRichTextValueFromUserList(userList, textStyleMap) {
	let userListString = '';

	const userTextStylesList = userList.map(({ emailAddress, role }) => {
		const startIndex = userListString.length;

		userListString += `${emailAddress} (${role})\n`;

		return {
			startIndex: startIndex,
			endIndex: userListString.length,
			textStyle: textStyleMap[role]
		};
	});

	// Remove trailing newline and replace with a space to prevent text style indices from being out of bounds
	const rtv = SpreadsheetApp.newRichTextValue().setText(userListString.slice(0, -1) + ' ');

	userTextStylesList.forEach(({ startIndex, endIndex, textStyle }) => {
		rtv.setTextStyle(startIndex, endIndex, textStyle);
	});

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
function createResultsSheet(sheetName, headerLabels, fileSummaryList, cellIconCache, colorMap) {
	console.log('Generating results sheet...');

	sheetName = `${sheetName} (${getDateString()})`;

	const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
	const numOfRows = fileSummaryList.length;
	const numOfCols = headerLabels.length;
	let resultsSheet = activeSpreadsheet.getSheetByName(sheetName);

	if (resultsSheet != null) {
		activeSpreadsheet.deleteSheet(resultsSheet);
	}

	resultsSheet = activeSpreadsheet.insertSheet(sheetName);

	resizeSheet(resultsSheet, numOfRows, numOfCols);

	resultsSheet.setFrozenRows(1);

	const headerRowRange = resultsSheet.getRange(1, 1, 1, numOfCols);
	const dataRange = resultsSheet.getRange(2, 1, numOfRows, numOfCols);
	const fileTypeDataRange = dataRange.offset(0, 1, numOfRows, 1);

	headerRowRange.setValues([headerLabels]);
	headerRowRange.setFontWeight('bold');
	dataRange.setVerticalAlignment('middle');
	fileTypeDataRange.setHorizontalAlignment('center');
	activeSpreadsheet.setActiveSheet(resultsSheet);

	// Transform the data and add it to the sheet
	const sortedFileSummaryList = fileSummaryList.sort(({ path: path1 }, { path: path2 }) => path1.localeCompare(path2));

	setRichTextValues(dataRange, sortedFileSummaryList, colorMap);

	// Set column widths
	resultsSheet.setColumnWidth(1, 50);
	resultsSheet.setColumnWidth(2, 20);
	resultsSheet.autoResizeColumns(3, numOfCols - 2);

	// Set the background color of the folder rows
	sortedFileSummaryList.forEach(({ isFolder }, i) => {
		dataRange.offset(i, 0, 1).setBackground(isFolder ? colorMap.folderBg : colorMap.fileBg);
	});

	setCellIconValues(fileTypeDataRange, sortedFileSummaryList, cellIconCache);

	console.log('Done generating results sheet');
}


/*
 * Map data to rich text values and assign them to the given range.
 */
function setRichTextValues(dataRange, sortedFileSummaryList, colorMap) {
	const textStyleMap = {
		reader: SpreadsheetApp.newTextStyle().setForegroundColor(colorMap.readerFont).build(),
		writer: SpreadsheetApp.newTextStyle().setForegroundColor(colorMap.writerFont).build(),
		commenter: SpreadsheetApp.newTextStyle().setForegroundColor(colorMap.commenterFont).build()
	};
	const blankRichTextValue = createRichTextValue('');
	const richTextValues = sortedFileSummaryList.map(fileSummary => [
		createRichTextValue(fileSummary.id),
		blankRichTextValue,
		createRichTextValue(fileSummary.path, fileSummary.link),
		createRichTextValueFromUserList(fileSummary.users, textStyleMap)
	]);

	dataRange.setRichTextValues(richTextValues);
}


/*
 * Map icon links to cell icons and assign them to the given range.
 */
function setCellIconValues(fileTypeDataRange, sortedFileSummaryList, cellIconCache) {
	const cellIconValues = sortedFileSummaryList.map(({ iconLink }) => {
		return cellIconCache.get(iconLink, () => [createCellImage(iconLink)]);
	});

	console.log('Loading file type icons. This may take a minute to complete...');

	fileTypeDataRange.setValues(cellIconValues);
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
 * Returns the current date as a string.
 */
function getDateString() {
	const date = new Date();

	return `${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`;
}
