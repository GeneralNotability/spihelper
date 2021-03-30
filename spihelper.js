/* eslint-disable es/no-object-entries */
/* eslint-disable no-restricted-syntax */
// <nowiki>
// @ts-check
// GeneralNotability's rewrite of Tim's SPI helper script
// v2.5.1 "Ignore all essays"

// Adapted from [[User:Mr.Z-man/closeAFD]]
importStylesheet('User:GeneralNotability/spihelper.css' );
importScript('User:Timotheus Canens/displaymessage.js');

// Typedefs
/**
 * @typedef SelectOption
 * @type {Object}
 * @property {string} label Text to display in the drop-down
 * @property {string} value Value to return if this option is selected
 * @property {boolean} selected Whether this item should be selected by default
 * @property {boolean=} disabled Whether this item should be disabled
 */

/**
 * @typedef BlockEntry
 * @type {Object}
 * @property {string} username Username to block
 * @property {string} duration Duration of block
 * @property {boolean} acb If set, account creation is blocked
 * @property {boolean} ab Whether autoblock is enabled (for registered users)/
 *     logged-in users are blocked (for IPs)
 * @property {boolean} ntp If set, talk page access is blocked
 * @property {boolean} nem If set, email access is blocked
 * @property {string} tpn Type of talk page notice to apply on block
 */

/**
 * @typedef TagEntry
 * @type {Object}
 * @property {string} username Username to tag
 * @property {string} tag Tag to apply to user
 * @property {string} altmasterTag Altmaster tag to apply to user, if relevant
 * @property {boolean} blocking Whether this account is marked for block as well
 */

 /**
  * @typedef ParsedArchiveNotice
  * @type {Object}
  * @property {string} username Case username
  * @property {boolean} xwiki Whether the crosswiki flag is set
  * @property {boolean} deny Whether the deny flag is set
  */

// Globals
// User-configurable settings, these are the defaults but will be updated by
// spiHelper_loadSettings()
const spiHelper_settings = {
	// Choices are 'watch' (unconditionally add to watchlist), 'preferences'
	// (follow default preferences), 'nochange' (don't change the watchlist
	// status of the page), and 'unwatch' (unconditionally remove)
	watchCase: 'preferences',
	watchCaseExpiry: 'indefinite',
	watchArchive: 'nochange',
	watchArchiveExpiry: 'indefinite',
	watchTaggedUser: 'preferences',
	watchTaggedUserExpiry: 'indefinite',
	watchNewCats: 'nochange',
	watchNewCatsExpiry: 'indefinite',
	watchBlockedUser: true,
	watchBlockedUserExpiry: 'indefinite',
	// Lets people disable clerk options if they're not a clerk
	clerk: true,
	// Log all actions to Special:MyPage/spihelper_log
	log: false,
	// Enable the "move section" button
	iUnderstandSectionMoves: false,
	// These are for debugging to view as other roles. If you're picking apart the code and
	// decide to set these (especially the CU option), it is YOUR responsibility to make sure
	// you don't do something that violates policy
	debugForceCheckuserState: null,
	debugForceAdminState: null
};

/** @type {string} Name of the SPI page in wiki title form
 * (e.g. Wikipedia:Sockpuppet investigations/Test) */
let spiHelper_pageName = mw.config.get('wgPageName').replace(/_/g, ' ');

/** @type {number} The main page's ID - used to check if the page
 * has been edited since we opened it to prevent edit conflicts
 */
let spiHelper_startingRevID = mw.config.get('wgCurRevisionId');

// Just the username part of the case
let spiHelper_caseName = spiHelper_pageName.replace(/Wikipedia:Sockpuppet investigations\//g, '');

/** list of section IDs + names corresponding to separate investigations */
let spiHelper_caseSections = [];

/** @type {?number} Selected section, "null" means that we're opearting on the entire page */
let spiHelper_sectionId = null;

/** @type {?string} Selected section's name (e.g. "10 June 2020") */
let spiHelper_sectionName = null;

/** @type {ParsedArchiveNotice} */
let spiHelper_archiveNoticeParams;

/** Map of top-level actions the user has selected */
const spiHelper_ActionsSelected = {
	Case_act: false,
	Block: false,
	Note: false,
	Close: false,
	Rename: false,
	Archive: false,
	SpiMgmt: false
};

/** @type {BlockEntry[]} Requested blocks */
const spiHelper_blocks = [];

/** @type {TagEntry[]} Requested tags */
const spiHelper_tags = [];

/** @type {string[]} Requested global locks */
const spiHelper_globalLocks = [];

// Count of unique users in the case (anything with a checkuser, checkip, user, ip, or vandal template on the page)
let spiHelper_usercount = 0;
const spiHelper_SECTION_RE = /^(?:===[^=]*===|=====[^=]*=====)\s*$/m;

/** @type {SelectOption[]} List of possible selections for tagging a user in the block/tag interface
 */
const spiHelper_TAG_OPTIONS = [
	{ label: 'None', selected: true, value: '' },
	{ label: 'Suspected sock', value: 'blocked', selected: false },
	{ label: 'Proven sock', value: 'proven', selected: false },
	{ label: 'CU confirmed sock', value: 'confirmed', selected: false },
	{ label: 'Blocked master', value: 'master', selected: false },
	{ label: 'CU confirmed master', value: 'sockmasterchecked', selected: false },
	{ label: '3X banned master', value: 'bannedmaster', selected: false }
];

/** @type {SelectOption[]} List of possible selections for tagging a user's altmaster in the block/tag interface */
const spiHelper_ALTMASTER_TAG_OPTIONS = [
	{ label: 'None', selected: true, value: '' },
	{ label: 'Suspected alt master', value: 'suspected', selected: false },
	{ label: 'Proven alt master', value: 'proven', selected: false }
];

/** @type {SelectOption[]} List of templates that CUs might insert */
const spiHelper_CU_TEMPLATES = [
	{ label: 'CU templates', selected: true, value: '', disabled: true },
	{ label: 'Confirmed', selected: false, value: '{{confirmed}}' },
	{ label: 'Confirmed/No Comment', selected: false, value: '{{confirmed-nc}}' },
	{ label: 'Indistinguishable', selected: false, value: '{{tallyho}}' },
	{ label: 'Likely', selected: false, value: '{{likely}}' },
	{ label: 'Possilikely', selected: false, value: '{{possilikely}}' },
	{ label: 'Possible', selected: false, value: '{{possible}}' },
	{ label: 'Unlikely', selected: false, value: '{{unlikely}}' },
	{ label: 'Unrelated', selected: false, value: '{{unrelated}}' },
	{ label: 'Inconclusive', selected: false, value: '{{inconclusive}}' },
	{ label: 'Need behavioral eval', selected: false, value: '{{behav}}' },
	{ label: 'No sleepers', selected: false, value: '{{nosleepers}}' },
	{ label: 'Stale', selected: false, value: '{{IPstale}}' },
	{ label: 'No comment (IP)', selected: false, value: '{{ncip}}'}
];

/** @type {SelectOption[]} Templates that a clerk or admin might insert */
const spiHelper_ADMIN_TEMPLATES = [
	{ label: 'Admin/clerk templates', selected: true, value: '', disabled: true },
	{ label: 'Duck', selected: false, value: '{{duck}}' },
	{ label: 'Megaphone Duck', selected: false, value: '{{megaphone duck}}' },
	{ label: 'Blocked and tagged', selected: false, value: '{{bnt}}' },
	{ label: 'Blocked, no tags', selected: false, value: '{{bwt}}' },
	{ label: 'Blocked, awaiting tags', selected: false, value: '{{sblock}}' },
	{ label: 'Blocked, tagged, closed', selected: false, value: '{{btc}}' },
	{ label: 'Diffs needed', selected: false, value: '{{DiffsNeeded|moreinfo}}' },
	{ label: 'Locks requested', selected: false, value: '{{GlobalLocksRequested}}' }
];

// Regex to match the case status, group 1 is the actual status
const spiHelper_CASESTATUS_RE = /{{\s*SPI case status\s*\|?\s*(\S*?)\s*}}/i;
// Regex to match closed case statuses (close or closed)
const spiHelper_CASESTATUS_CLOSED_RE = /^closed?$/i;

const spiHelper_CLERKSTATUS_RE = /{{(CURequest|awaitingadmin|clerk ?request|(?:self|requestand|cu)?endorse|inprogress|decline(?:-ip)?|moreinfo|relisted|onhold)}}/i;

const spiHelper_SOCK_SECTION_RE_WITH_NEWLINE = /====\s*Suspected sockpuppets\s*====\n*/i;

const spiHelper_ADMIN_SECTION_RE = /\s*====\s*<big>Clerk, CheckUser, and\/or patrolling admin comments<\/big>\s*====\s*/i;

const spiHelper_CU_BLOCK_RE = /{{(checkuserblock(-account|-wide)?|checkuser block)}}/i;

const spiHelper_ARCHIVENOTICE_RE = /{{\s*SPI\s*archive notice\|(?:1=)?([^|]*?)(\|.*)?}}/i;

const spiHelper_PRIORCASES_RE = /{{spipriorcases}}/i;

// regex to remove hidden characters from form inputs - they mess up some things,
// especially mw.util.isIP
const spiHelper_HIDDEN_CHAR_NORM_RE = /\u200E/;

const spihelper_ADVERT = ' (using [[:w:en:User:GeneralNotability/spihelper|spihelper.js]])';

// The current wiki's interwiki prefix
const spiHelper_interwikiPrefix = spiHelper_getInterwikiPrefix();

// Map of active operations (used as a "dirty" flag for beforeunload)
// Values are strings representing the state - acceptable values are 'running', 'success', 'failed'
const spiHelper_activeOperations = new Map();

// Actually put the portlets in place if needed
if (mw.config.get('wgPageName').includes('Wikipedia:Sockpuppet_investigations/') &&
	!mw.config.get('wgPageName').includes('Wikipedia:Sockpuppet_investigations/SPI/') &&
	!mw.config.get('wgPageName').includes('/Archive')) {
	mw.loader.load('mediawiki.user');
	$(spiHelper_addLink);
}

// Main functions - do the meat of the processing and UI work

const spiHelper_TOP_VIEW = `
<div id="spiHelper_topViewDiv">
	<h3>Handling SPI case</h3>
	<select id="spiHelper_sectionSelect"/>
	<ul>
		<li id="spiHelper_actionLine"  class="spiHelper_singleCaseOnly">
			<input type="checkbox" name="spiHelper_Case_Action" id="spiHelper_Case_Action" />
			<label for="spiHelper_Case_Action">Change case status</label>
		</li>
		<li id="spiHelper_spiMgmtLine"  class="spiHelper_allCasesOnly">
			<input type="checkbox" id="spiHelper_SpiMgmt" />
			<label for="spiHelper_SpiMgmt">Change SPI options</label>
		</li>
		<li id="spiHelper_blockLine" class="spiHelper_adminClerkClass">
			<input type="checkbox" name="spiHelper_BlockTag" id="spiHelper_BlockTag" />
			<label for="spiHelper_BlockTag">Block/tag socks</label>
		</li>
		<li id="spiHelper_commentLine" class="spiHelper_singleCaseOnly">
			<input type="checkbox" name="spiHelper_Comment" id="spiHelper_Comment" />
			<label for="spiHelper_Comment">Note/comment</label>
  		</li>
		<li id="spiHelper_closeLine" class="spiHelper_adminClerkClass spiHelper_singleCaseOnly">
			<input type="checkbox" name="spiHelper_Close" id="spiHelper_Close")" />
			<label for="spiHelper_Close">Close case</label>
		</li>
		<li id="spiHelper_moveLine" class="spiHelper_clerkClass">
			<input type="checkbox" name="spiHelper_Move" id="spiHelper_Move" />
			<label for="spiHelper_Move" id="spiHelper_moveLabel">Move/merge full case (Clerk only)</label>
		</li>
		<li id="spiHelper_archiveLine" class="spiHelper_clerkClass">
			<input type="checkbox" name="spiHelper_Archive" id="spiHelper_Archive"/>
			<label for="spiHelper_Archive">Archive case (Clerk only)</label>
		</li>
	</ul>
	<input type="button" id="spiHelper_GenerateForm" name="spiHelper_GenerateForm" value="Continue" onclick="spiHelper_generateForm()" />
</div>
`;

/**
 * Initialization functions for spiHelper, displays the top-level menu
 */
async function spiHelper_init() {
	'use strict';
	spiHelper_caseSections = await spiHelper_getInvestigationSectionIDs();
	
	// Load archivenotice params
	spiHelper_archiveNoticeParams = await spiHelper_parseArchiveNotice(spiHelper_pageName);

	// First, insert the template text
	displayMessage(spiHelper_TOP_VIEW);

	// Narrow search scope
	const $topView = $('#spiHelper_topViewDiv', document);

	// Next, modify what's displayed
	// Set the block selection label based on whether or not the user is an admin
	$('#spiHelper_blockLabel', $topView).text(spiHelper_isAdmin() ? 'Block/tag socks' : 'Tag socks');

	// Wire up a couple of onclick handlers
	$('#spiHelper_Move', $topView).on('click', function () {
		spiHelper_updateArchive();
	});
	$('#spiHelper_Archive', $topView).on('click', function () {
		spiHelper_updateMove();
	});

	// Generate the section selector
	const $sectionSelect = $('#spiHelper_sectionSelect', $topView);
	$sectionSelect.on('change', () => {
		spiHelper_setCheckboxesBySection();
	});

	// Add the dates to the selector
	for (let i = 0; i < spiHelper_caseSections.length; i++) {
		const s = spiHelper_caseSections[i];
		$('<option>').val(s.index).text(s.line).appendTo($sectionSelect);
	}
	// All-sections selector...deliberately at the bottom, the default should be the first section
	$('<option>').val('all').text('All Sections').appendTo($sectionSelect);

	// Hide block and close from non-admin non-clerks
	if (!(spiHelper_isAdmin() || spiHelper_isClerk())) {
		$('.spiHelper_adminClerkClass', $topView).hide();
	}

	// Hide move and archive from non-clerks
	if (!spiHelper_isClerk()) {
		$('.spiHelper_clerkClass', $topView).hide();
	}

	// Set the checkboxes to their default states
	spiHelper_setCheckboxesBySection();
}

const spiHelper_ACTION_VIEW = `
<div id="spiHelper_actionViewDiv">
	<small><a id="spiHelper_backLink">Back to top menu</a></small>
	<br />
	<h3>Handling SPI case</h3>
	<div id="spiHelper_actionView">
		<h4>Changing case status</h4>
		<label for="spiHelper_CaseAction">New status:</label>
		<select id="spiHelper_CaseAction"/>
	</div>
	<div id="spiHelper_spiMgmtView">
		<h4>Changing SPI settings</h4>
		<ul>
			<li>
				<input type="checkbox" id="spiHelper_spiMgmt_crosswiki" />
				<label for="spiHelper_Case_Action">Case is crosswiki</label>
			</li>
			<li>
				<input type="checkbox" id="spiHelper_spiMgmt_deny" />
				<label for="spiHelper_Case_Action">Socks should not be tagged per DENY</label>
			</li>
		</ul>
	</div>
	<div id="spiHelper_blockTagView">
		<h4 id="spiHelper_blockTagHeader">Blocking and tagging socks</h4>
		<ul>
			<li class="spiHelper_adminClass">
				<input type="checkbox" name="spiHelper_noblock" id="spiHelper_noblock" />
				<label for="spiHelper_noblock">Do not make any blocks (this overrides the individual "Blk" boxes below)</label>
			</li>
			<li class="spiHelper_adminClass">
				<input type="checkbox" name="spiHelper_override" id="spiHelper_override" />
				<label for="spiHelper_override">Override any existing blocks</label>
			</li>
			<li class="spiHelper_cuClass">
				<input type="checkbox" name="spiHelper_cublock" id="spiHelper_cublock" />
				<label for="spiHelper_cublock">Mark blocks as Checkuser blocks.</label>
			</li>
			<li class="spiHelper_cuClass">
				<input type="checkbox" name="spiHelper_cublockonly" id="spiHelper_cublockonly" />
				<label for="spiHelper_cublockonly">
					Suppress the usual block summary and only use {{checkuserblock-account}} and {{checkuserblock}} (no effect if "mark blocks as CU blocks" is not checked).
				</label>
			</li>
			<li class="spiHelper_adminClass">
				<input type="checkbox" checked="checked" name="spiHelper_blocknoticemaster" id="spiHelper_blocknoticemaster" />
				<label for="spiHelper_blocknoticemaster">Add talk page notice when (re)blocking the sockmaster.</label>
			</li>
			<li class="spiHelper_adminClass">
				<input type="checkbox" checked="checked" name="spiHelper_blocknoticesocks" id="spiHelper_blocknoticesocks" />
				<label for="spiHelper_blocknoticesocks">Add talk page notice when blocking socks.</label>
			</li>
			<li class="spiHelper_adminClass">
				<input type="checkbox" name="spiHelper_blanktalk" id="spiHelper_blanktalk" />
				<label for="spiHelper_blanktalk">Blank the talk page when adding talk notices.</label>
			</li>
			<li>
				<input type="checkbox" name="spiHelper_hidelocknames" id="spiHelper_hidelocknames" />
				<label for="spiHelper_hidelocknames">Hide usernames when requesting global locks.</label>
			</li>
		</ul>
		<table id="spiHelper_blockTable" style="border-collapse:collapse;">
			<tr>
				<th>Username</th>
				<th class="spiHelper_adminClass"><span title="Block user" class="rt-commentedText spihelper-hovertext">Blk?</span></th>
				<th class="spiHelper_adminClass"><span title="Block duration" class="rt-commentedText spihelper-hovertext">Duration</span></th>
				<th class="spiHelper_adminClass"><span title="Account creation blocked" class="rt-commentedText spihelper-hovertext">ACB</span></th>
				<th class="spiHelper_adminClass"><span title="Autoblock (for logged-in users)/Anonymous-only (for IPs)" class="rt-commentedText spihelper-hovertext">AB/AO</span></th>
				<th class="spiHelper_adminClass"><span title="Disable talk page access" class="rt-commentedText spihelper-hovertext">NTP</span></th>
				<th class="spiHelper_adminClass"><span title="Disable email" class="rt-commentedText spihelper-hovertext">NEM</span></th>
				<th>Tag</th>
				<th><span title="Tag the user with a suspected alternate master" class="rt-commentedText spihelper-hovertext">Alt Master</span></th>
				<th><span title="Request a global lock at Meta:SRG" class="rt-commentedText spihelper-hovertext">Req Lock?</span></th>
			</tr>
			<tr style="border-bottom:2px solid black">
				<td style="text-align:center;">(All users)</td>
				<td class="spiHelper_adminClass"><input type="checkbox" id="spiHelper_block_doblock"/></td>
				<td class="spiHelper_adminClass"></td>
				<td class="spiHelper_adminClass"><input type="checkbox" id="spiHelper_block_acb" checked="checked"/></td>
				<td class="spiHelper_adminClass"><input type="checkbox" id="spiHelper_block_ab" checked="checked"/></td>
				<td class="spiHelper_adminClass"><input type="checkbox" id="spiHelper_block_tp"/></td>
				<td class="spiHelper_adminClass"><input type="checkbox" id="spiHelper_block_email"/></td>
				<td><select id="spiHelper_block_tag"/></td>
				<td><select id="spiHelper_block_tag_altmaster"/></td>
	
				<td><input type="checkbox" name="spiHelper_block_lock_all" id="spiHelper_block_lock"/></td>
			</tr>
		</table>
	</div>
	<div id="spiHelper_closeView">
		<h4>Marking case as closed</h4>
		<input type="checkbox" checked="checked" id="spiHelper_CloseCase" />
		<label for="spiHelper_CloseCase">Close this SPI case</label>
	</div>
	<div id="spiHelper_moveView">
		<h4 id="spiHelper_moveHeader">Move section</h4>
		<label for="spiHelper_moveTarget">New sockmaster username: </label>
		<input type="text" name="spiHelper_moveTarget" id="spiHelper_moveTarget" />
	</div>
	<div id="spiHelper_archiveView">
		<h4>Archiving case</h4>
		<input type="checkbox" checked="checked" name="spiHelper_ArchiveCase" id="spiHelper_ArchiveCase" />
		<label for="spiHelper_ArchiveCase">Archive this SPI case</label>
	</div>
	<div id="spiHelper_commentView">
		<h4>Comments</h4>
		<span>
			<select id="spiHelper_noteSelect"/>
			<select class="spiHelper_adminClerkClass" id="spiHelper_adminSelect"/>
			<select class="spiHelper_cuClass" id="spiHelper_cuSelect"/>
		</span>
		<div>
			<label for="spiHelper_CommentText">Comment:</label>
			<textarea rows="3" cols="80" id="spiHelper_CommentText">*</textarea>
			<div><a id="spiHelper_previewLink">Preview</a></div>
		</div>
		<div class="spihelper-previewbox" id="spiHelper_previewBox" hidden/>
	</div>
	<input type="button" id="spiHelper_performActions" value="Done" />
</div>
`;
/**
 * Big function to generate the SPI form from the top-level menu selections
 */
async function spiHelper_generateForm() {
	'use strict';
	spiHelper_usercount = 0;
	const $topView = $('#spiHelper_topViewDiv', document);
	spiHelper_ActionsSelected.Case_act = $('#spiHelper_Case_Action', $topView).prop('checked');
	spiHelper_ActionsSelected.Block = $('#spiHelper_BlockTag', $topView).prop('checked');
	spiHelper_ActionsSelected.Note = $('#spiHelper_Comment', $topView).prop('checked');
	spiHelper_ActionsSelected.Close = $('#spiHelper_Close', $topView).prop('checked');
	spiHelper_ActionsSelected.Rename = $('#spiHelper_Move', $topView).prop('checked');
	spiHelper_ActionsSelected.Archive = $('#spiHelper_Archive', $topView).prop('checked');
	spiHelper_ActionsSelected.SpiMgmt = $('#spiHelper_SpiMgmt', $topView).prop('checked');
	const pagetext = await spiHelper_getPageText(spiHelper_pageName, false, spiHelper_sectionId);
	if (!(spiHelper_ActionsSelected.Case_act ||
		spiHelper_ActionsSelected.Note || spiHelper_ActionsSelected.Close ||
		spiHelper_ActionsSelected.Archive || spiHelper_ActionsSelected.Block ||
		spiHelper_ActionsSelected.Rename || spiHelper_ActionsSelected.SpiMgmt)) {
		displayMessage('');
		return;
	}

	displayMessage(spiHelper_ACTION_VIEW);

	// Reduce the scope that jquery operates on
	const $actionView = $('#spiHelper_actionViewDiv', document);

	// Wire up the action view
	$('#spiHelper_backLink', $actionView).on('click', () => {
		spiHelper_init();
	});
	if (spiHelper_ActionsSelected.Case_act) {
		const result = spiHelper_CASESTATUS_RE.exec(pagetext);
		let casestatus = '';
		if (result) {
			casestatus = result[1];
		}
		const canAddCURequest = (casestatus === '' || /^(?:admin|moreinfo|cumoreinfo|hold|cuhold|clerk|open)$/i.test(casestatus));
		const cuRequested = /^(?:CU|checkuser|CUrequest|request|cumoreinfo)$/i.test(casestatus);
		const cuEndorsed = /^(?:endorse(d)?)$/i.test(casestatus);
		const cuCompleted = /^(?:inprogress|checking|relist(ed)?|checked|completed|declined?|cudeclin(ed)?)$/i.test(casestatus);

		/** @type {SelectOption[]} Generated array of values for the case status select box */
		const selectOpts = [
			{ label: 'No action', value: 'noaction', selected: true }
		];
		if (spiHelper_CASESTATUS_CLOSED_RE.test(casestatus)) {
			selectOpts.push({ label: 'Reopen', value: 'open', selected: false });
		}
		if (spiHelper_isCheckuser()) {
			selectOpts.push({ label: 'Mark as in progress', value: 'inprogress', selected: false });
		}
		if (spiHelper_isClerk() || spiHelper_isAdmin()) {
			selectOpts.push({ label: 'Request more information', value: 'moreinfo', selected: false });
		}
		if (canAddCURequest) {
			// Statuses only available if the case could be moved to "CU requested"
			selectOpts.push({ label: 'Request CU', value: 'CUrequest', selected: false });
			if (spiHelper_isClerk()) {
				selectOpts.push({ label: 'Request CU and self-endorse', value: 'selfendorse', selected: false });
			}
		}
		// CU already requested
		if (cuRequested && spiHelper_isClerk()) {
			// Statuses only available if CU has been requested, only clerks + CUs should use these
			selectOpts.push({ label: 'Endorse for CU attention', value: 'endorse', selected: false });
			// Switch the decline option depending on whether the user is a checkuser
			if (spiHelper_isCheckuser()) {
				selectOpts.push({ label: 'Endorse CU as a CheckUser', value: 'cuendorse', selected: false });
			}
			if (spiHelper_isCheckuser()) {
				selectOpts.push({ label: 'Decline CU', value: 'cudecline', selected: false });
			}
			else {
				selectOpts.push({ label: 'Decline CU', value: 'decline', selected: false });
			}
			selectOpts.push({ label: 'Request more information for CU', value: 'cumoreinfo', selected: false });
		} else if (cuEndorsed && spiHelper_isCheckuser()) {
			// Let checkusers decline endorsed cases
			if (spiHelper_isCheckuser()) {
				selectOpts.push({ label: 'Decline CU', value: 'cudecline', selected: false });
			}
			selectOpts.push({ label: 'Request more information for CU', value: 'cumoreinfo', selected: false });
		}
		// This is mostly a CU function, but let's let clerks and admins set it
		//  in case the CU forgot (or in case we're un-closing))
		if (spiHelper_isAdmin() || spiHelper_isClerk()) {
			selectOpts.push({ label: 'Mark as checked', value: 'checked', selected: false });
		}
		if (spiHelper_isClerk() && cuCompleted) {
			selectOpts.push({ label: 'Relist for another check', value: 'relist', selected: false });
		}
		if (spiHelper_isCheckuser()) {
			selectOpts.push({ label: 'Place case on CU hold', value: 'cuhold', selected: false });
		} else { // I guess it's okay for anyone to have this option
			selectOpts.push({ label: 'Place case on hold', value: 'hold', selected: false });
		}
		selectOpts.push({ label: 'Request clerk action', value: 'clerk', selected: false });
		// I think this is only useful for non-admin clerks to ask admins to do stuff
		if (!spiHelper_isAdmin() && spiHelper_isClerk()) {
			selectOpts.push({ label: 'Request admin action', value: 'admin', selected: false });
		}
		// Generate the case action options
		spiHelper_generateSelect('spiHelper_CaseAction', selectOpts);
		// Add the onclick handler to the drop-down
		$('#spiHelper_CaseAction', $actionView).on('change', function (e) {
			spiHelper_caseActionUpdated($(e.target));
		});
	} else {
		$('#spiHelper_actionView', $actionView).hide();
	}

	if (spiHelper_ActionsSelected.SpiMgmt) {
		const $xwikiBox = $('#spiHelper_spiMgmt_crosswiki', $actionView);
		const $denyBox = $('#spiHelper_spiMgmt_deny', $actionView);

		$xwikiBox.prop('checked', spiHelper_archiveNoticeParams.xwiki);
		$denyBox.prop('checked', spiHelper_archiveNoticeParams.deny);
	} else {
		$('#spiHelper_spiMgmtView', $actionView).hide();
	}

	if (spiHelper_ActionsSelected.Block) {
		if (spiHelper_isAdmin()) {
			$('#spiHelper_blockTagHeader', $actionView).text('Blocking and tagging socks');
		} else {
			$('#spiHelper_blockTagHeader', $actionView).text('Tagging socks');
		}
		const checkuser_re = /{{\s*check(?:user|ip)\s*\|\s*(?:1=)?\s*([^\|}]*?)\s*(?:\|master name\s*=\s*.*)?}}/gi;
		const results = pagetext.match(checkuser_re);
		const likelyusers = [];
		const likelyips = [];
		const possibleusers = [];
		const possibleips = [];
		likelyusers.push(spiHelper_caseName);
		if (results) {
			for (let i = 0; i < results.length; i++) {
				const username = spiHelper_normalizeUsername(results[i].replace(checkuser_re, '$1'));
				const isIP = mw.util.isIPAddress(username, true);
				if (!isIP && !likelyusers.includes(username)) {
					likelyusers.push(username);
				} else if (isIP && !likelyips.includes(username)) {
					likelyips.push(username);
				}
			}
		}
		const user_re = /{{\s*(?:user|vandal|IP)[^\|}{]*?\s*\|\s*(?:1=)?\s*([^\|}]*?)\s*}}/gi;
		const userresults = pagetext.match(user_re);
		if (userresults) {
			for (let i = 0; i < userresults.length; i++) {
				const username = spiHelper_normalizeUsername(userresults[i].replace(user_re, '$1'));
				if (mw.util.isIPAddress(username, true) && !possibleips.includes(username) &&
					!likelyips.includes(username)) {
					possibleips.push(username);
				} else if (!possibleusers.includes(username) &&
					!likelyusers.includes(username)) {
					possibleusers.push(username);
				}
			}
		}
		// Wire up the "select all" options
		$('#spiHelper_block_doblock', $actionView).on('click', function (e) {
			spiHelper_setAllBlockOpts($(e.target));
		});
		$('#spiHelper_block_acb', $actionView).on('click', function (e) {
			spiHelper_setAllBlockOpts($(e.target));
		});
		$('#spiHelper_block_ab', $actionView).on('click', function (e) {
			spiHelper_setAllBlockOpts($(e.target));
		});
		$('#spiHelper_block_tp', $actionView).on('click', function (e) {
			spiHelper_setAllBlockOpts($(e.target));
		});
		$('#spiHelper_block_email', $actionView).on('click', function (e) {
			spiHelper_setAllBlockOpts($(e.target));
		});
		$('#spiHelper_block_lock', $actionView).on('click', function (e) {
			spiHelper_setAllBlockOpts($(e.target));
		});
		$('#spiHelper_block_lock', $actionView).on('click', function (e) {
			spiHelper_setAllBlockOpts($(e.target));
		});
		spiHelper_generateSelect('spiHelper_block_tag', spiHelper_TAG_OPTIONS);
		$('#spiHelper_block_tag', $actionView).on('change', function (e) {
			spiHelper_setAllBlockOpts($(e.target));
		});
		spiHelper_generateSelect('spiHelper_block_tag_altmaster', spiHelper_ALTMASTER_TAG_OPTIONS);
		$('#spiHelper_block_tag_altmaster', $actionView).on('change', function (e) {
			spiHelper_setAllBlockOpts($(e.target));
		});
		$('#spiHelper_block_lock', $actionView).on('click', function (e) {
			spiHelper_setAllBlockOpts($(e.target));
		});

		for (let i = 0; i < likelyusers.length; i++) {
			spiHelper_usercount++;
			spiHelper_generateBlockTableLine(likelyusers[i], true, spiHelper_usercount);
		}
		for (let i = 0; i < likelyips.length; i++) {
			spiHelper_usercount++;
			spiHelper_generateBlockTableLine(likelyips[i], true, spiHelper_usercount);
		}
		for (let i = 0; i < possibleusers.length; i++) {
			spiHelper_usercount++;
			spiHelper_generateBlockTableLine(possibleusers[i], false, spiHelper_usercount);
		}
		for (let i = 0; i < possibleips.length; i++) {
			spiHelper_usercount++;
			spiHelper_generateBlockTableLine(possibleips[i], false, spiHelper_usercount);
		}
	} else {
		$('#spiHelper_blockTagView', $actionView).hide();
	}
	if (!spiHelper_ActionsSelected.Close) {
		$('#spiHelper_closeView', $actionView).hide();
	}
	if (spiHelper_ActionsSelected.Rename) {
		if (spiHelper_sectionId) {
			$('#spiHelper_moveHeader', $actionView).text('Move section "' + spiHelper_sectionName + '"');
		} else {
			$('#spiHelper_moveHeader', $actionView).text('Move/merge full case');

		}
	} else {
		$('#spiHelper_moveView', $actionView).hide();
	}

	if (!spiHelper_ActionsSelected.Archive) {
		$('#spiHelper_archiveView', $actionView).hide();
	}

	// Only give the option to comment if we selected a specific section
	if (spiHelper_sectionId) {
		// generate the note prefixes
		/** @type {SelectOption[]} */
		const spiHelper_noteTemplates = [
			{ label: 'Comment templates', selected: true, value: '', disabled: true }
		];
		if (spiHelper_isClerk()) {
			spiHelper_noteTemplates.push({ label: 'Clerk note', selected: false, value: 'clerknote' });
		}
		if (spiHelper_isAdmin()) {
			spiHelper_noteTemplates.push({ label: 'Administrator note', selected: false, value: 'adminnote' });
		}
		if (spiHelper_isCheckuser()) {
			spiHelper_noteTemplates.push({ label: 'CU note', selected: false, value: 'cunote' });
		}
		spiHelper_noteTemplates.push({ label: 'Note', selected: false, value: 'takenote' });

		// Wire up the select boxes
		spiHelper_generateSelect('spiHelper_noteSelect', spiHelper_noteTemplates);
		$('#spiHelper_noteSelect', $actionView).on('change', function (e) {
			spiHelper_insertNote($(e.target));
		});
		spiHelper_generateSelect('spiHelper_adminSelect', spiHelper_ADMIN_TEMPLATES);
		$('#spiHelper_adminSelect', $actionView).on('change', function (e) {
			spiHelper_insertTextFromSelect($(e.target));
		});
		spiHelper_generateSelect('spiHelper_cuSelect', spiHelper_CU_TEMPLATES);
		$('#spiHelper_cuSelect', $actionView).on('change', function (e) {
			spiHelper_insertTextFromSelect($(e.target));
		});
		$('#spiHelper_previewLink', $actionView).on('click', function () {
			spiHelper_previewText();
		});
	} else {
		$('#spiHelper_commentView', $actionView).hide();
	}
	// Wire up the submit button
	$('#spiHelper_performActions', $actionView).on('click', () => {
		spiHelper_performActions();
	});

	// Hide items based on role
	if (!spiHelper_isCheckuser()) {
		// Hide CU options from non-CUs
		$('.spiHelper_cuClass', $actionView).hide();
	}
	if (!spiHelper_isAdmin()) {
		// Hide block options from non-admins
		$('.spiHelper_adminClass', $actionView).hide();
	}
	if (!(spiHelper_isAdmin() || spiHelper_isClerk())) {
		$('.spiHelper_adminClerkClass', $actionView).hide();
	}
}

/**
 * Archives everything on the page that's eligible for archiving
 */
async function spiHelper_oneClickArchive() {
	'use strict';
	spiHelper_activeOperations.set('oneClickArchive', 'running');

	const pagetext = await spiHelper_getPageText(spiHelper_pageName, false);
	spiHelper_caseSections = await spiHelper_getInvestigationSectionIDs();
	if (!spiHelper_SECTION_RE.test(pagetext)) {
		alert('Looks like the page has been archived already.');
		spiHelper_activeOperations.set('oneClickArchive', 'successful');
		return;
	}
	displayMessage('<ul id="spiHelper_status"/>');
	await spiHelper_archiveCase();
	await spiHelper_purgePage(spiHelper_pageName);
	const logMessage = '* [[' + spiHelper_pageName + ']]: used one-click archiver ~~~~~';
	if (spiHelper_settings.log) {
		spiHelper_log(logMessage);
	}
	$('#spiHelper_status', document).append($('<li>').text('Done!'));
	spiHelper_activeOperations.set('oneClickArchive', 'successful');
}

/**
 * Another "meaty" function - goes through the action selections and executes them
 */
async function spiHelper_performActions() {
	'use strict';
	spiHelper_activeOperations.set('mainActions', 'running');

	// Again, reduce the search scope
	const $actionView = $('#spiHelper_actionViewDiv', document);

	// set up a few function-scoped vars
	let comment = '';
	let cuBlock = false;
	let cuBlockOnly = false;
	let newCaseStatus = 'noaction';
	let renameTarget = '';

	/** @type {boolean} */
	const blankTalk = $('#spiHelper_blanktalk', $actionView).prop('checked');
	/** @type {boolean} */
	const overrideExisting = $('#spiHelper_override', $actionView).prop('checked');
	/** @type {boolean} */
	const hideLockNames = $('#spiHelper_hidelocknames', $actionView).prop('checked');

	if (spiHelper_ActionsSelected.Case_act) {
		newCaseStatus = $('#spiHelper_CaseAction', $actionView).val().toString();
	}
	if (spiHelper_ActionsSelected.SpiMgmt) {
		spiHelper_archiveNoticeParams.deny = $('#spiHelper_spiMgmt_deny', $actionView).prop('checked');
		spiHelper_archiveNoticeParams.xwiki = $('#spiHelper_spiMgmt_crosswiki', $actionView).prop('checked');
	}
	if (spiHelper_sectionId) {
		comment = $('#spiHelper_CommentText', $actionView).val().toString();
	}
	if (spiHelper_ActionsSelected.Block) {
		if (spiHelper_isCheckuser()) {
			cuBlock = $('#spiHelper_cublock', $actionView).prop('checked');
			cuBlockOnly = $('#spiHelper_cublockonly', $actionView).prop('checked');
		}
		if (spiHelper_isAdmin() && !$('#spiHelper_noblock', $actionView).prop('checked')) {
			const masterNotice = $('#spiHelper_blocknoticemaster', $actionView).prop('checked');
			const sockNotice = $('#spiHelper_blocknoticesocks', $actionView).prop('checked');
			for (let i = 1; i <= spiHelper_usercount; i++) {
				if ($('#spiHelper_block_doblock' + i, $actionView).prop('checked')) {
					let noticetype = '';

					if (masterNotice && $('#spiHelper_block_tag' + i, $actionView).val().toString().includes('master')) {
						noticetype = 'master';
					} else if (sockNotice) {
						noticetype = 'sock';
					}

					/** @type {BlockEntry} */
					const item = {
						username: spiHelper_normalizeUsername($('#spiHelper_block_username' + i, $actionView).val().toString()),
						duration: $('#spiHelper_block_duration' + i, $actionView).val().toString(),
						acb: $('#spiHelper_block_acb' + i, $actionView).prop('checked'),
						ab: $('#spiHelper_block_ab' + i, $actionView).prop('checked'),
						ntp: $('#spiHelper_block_tp' + i, $actionView).prop('checked'),
						nem: $('#spiHelper_block_email' + i, $actionView).prop('checked'),
						tpn: noticetype
					};

					spiHelper_blocks.push(item);
				}
				if ($('#spiHelper_block_lock' + i, $actionView).prop('checked')) {
					spiHelper_globalLocks.push($('#spiHelper_block_username' + i, $actionView).val().toString());
				}
				if ($('#spiHelper_block_tag' + i).val() !== '') {
					const item = {
						username: spiHelper_normalizeUsername($('#spiHelper_block_username' + i, $actionView).val().toString()),
						tag: $('#spiHelper_block_tag' + i, $actionView).val().toString(),
						altmasterTag: $('#spiHelper_block_tag_altmaster' + i, $actionView).val().toString(),
						blocking: $('#spiHelper_block_doblock' + i, $actionView).prop('checked')
					};
					spiHelper_tags.push(item);
				}
			}
		} else {
			for (let i = 1; i <= spiHelper_usercount; i++) {
				if ($('#spiHelper_block_tag' + i, $actionView).val() !== '') {
					const item = {
						username: spiHelper_normalizeUsername($('#spiHelper_block_username' + i, $actionView).val().toString()),
						tag: $('#spiHelper_block_tag' + i, $actionView).val().toString(),
						altmasterTag: $('#spiHelper_block_tag_altmaster' + i, $actionView).val().toString(),
						blocking: false
					};
					spiHelper_tags.push(item);
				}
				if ($('#spiHelper_block_lock' + i, $actionView).prop('checked')) {
					spiHelper_globalLocks.push(spiHelper_normalizeUsername($('#spiHelper_block_username' + i, $actionView).val().toString()));
				}
			}
		}
	}
	if (spiHelper_ActionsSelected.Close) {
		spiHelper_ActionsSelected.Close = $('#spiHelper_CloseCase', $actionView).prop('checked');
	}
	if (spiHelper_ActionsSelected.Rename) {
		renameTarget = spiHelper_normalizeUsername($('#spiHelper_moveTarget', $actionView).val().toString());
	}
	if (spiHelper_ActionsSelected.Archive) {
		spiHelper_ActionsSelected.Archive = $('#spiHelper_ArchiveCase', $actionView).prop('checked');
	}

	displayMessage('<ul id="spiHelper_status" />');

	const $statusAnchor = $('#spiHelper_status', document);

	let sectionText = await spiHelper_getPageText(spiHelper_pageName, true, spiHelper_sectionId);
	let editsummary = '';
	let logMessage = '* [[' + spiHelper_pageName + ']]';
	if (spiHelper_sectionId) {
		logMessage += ' (section ' + spiHelper_sectionName + ')';
	} else {
		logMessage += ' (full case)';
	}
	logMessage += ' ~~~~~';

	if (spiHelper_sectionId !== null) {
		let caseStatusResult = spiHelper_CASESTATUS_RE.exec(sectionText);
		if (caseStatusResult === null) {
			sectionText = sectionText.replace('===', '{{SPI case status|}}\n===');
			caseStatusResult = spiHelper_CASESTATUS_RE.exec(sectionText);
		}
		const oldCaseStatus = caseStatusResult[1] || 'open';
		if (newCaseStatus === 'noaction') {
			newCaseStatus = oldCaseStatus;
		}

		if (spiHelper_ActionsSelected.Case_act && newCaseStatus !== 'noaction') {
			switch (newCaseStatus) {
				case 'open':
					editsummary = 'Reopening';
					break;
				case 'CUrequest':
					editsummary = 'Adding checkuser request';
					break;
				case 'admin':
					editsummary = 'Requesting admin action';
					break;
				case 'clerk':
					editsummary = 'Requesting clerk action';
					break;
				case 'selfendorse':
					newCaseStatus = 'endorse';
					editsummary = 'Adding checkuser request (self-endorsed for checkuser attention)';
					break;
				case 'checked':
					editsummary = 'Marking request as checked';
					break;
				case 'inprogress':
					editsummary = 'Marking request in progress';
					break;
				case 'decline':
					editsummary = 'Declining checkuser';
					break;
				case 'cudecline':
					editsummary = 'CU declining checkuser';
					break;
				case 'endorse':
					editsummary = 'Endorsing for checkuser attention';
					break;
				case 'cuendorse':
					editsummary = 'CU endorsing for checkuser attention';
					break;
				case 'moreinfo': // Intentional fallthrough
				case 'cumoreinfo':
					editsummary = 'Requesting additional information';
					break;
				case 'relist':
					editsummary = 'Relisting case for another check';
					break;
				case 'hold':
					editsummary = 'Putting case on hold';
					break;
				case 'cuhold':
					editsummary = 'Placing checkuser request on hold';
					break;
				case 'noaction':
					// Do nothing
					break;
				default:
					console.error('Unexpected case status value ' + newCaseStatus);
			}
			logMessage += '\n** changed case status from ' + oldCaseStatus + ' to ' + newCaseStatus;
		}
	}

	if (spiHelper_ActionsSelected.SpiMgmt) {
		let newArchiveNotice = spiHelper_makeNewArchiveNotice(spiHelper_caseName, spiHelper_archiveNoticeParams);
		sectionText = sectionText.replace(spiHelper_ARCHIVENOTICE_RE, newArchiveNotice);
		if (editsummary) {
			editsummary += ', update archivenotice';
		} else {
			editsummary = 'Update archivenotice';
		}
		logMessage += '\n** Updated archivenotice';
	}

	if (spiHelper_ActionsSelected.Block) {
		let sockmaster = '';
		let altmaster = '';
		let needsAltmaster = false;
		spiHelper_tags.forEach(async (tagEntry) => {
			// we do not support tagging IPs
			if (mw.util.isIPAddress(tagEntry.username, true)) {
				// Skip, this is an IP
				return;
			}
			if (tagEntry.tag.includes('master')) {
				sockmaster = tagEntry.username;
			}
			if (tagEntry.altmasterTag !== '') {
				needsAltmaster = true;
			}
		});
		if (sockmaster === '') {
			sockmaster = prompt('Please enter the name of the sockmaster: ', spiHelper_caseName);
		}
		if (needsAltmaster) {
			altmaster = prompt('Please enter the name of the alternate sockmaster: ', spiHelper_caseName);
		}

		let blockedList = '';
		if (spiHelper_isAdmin()) {
			spiHelper_blocks.forEach(async (blockEntry) => {
				const blockReason = await spiHelper_getUserBlockReason(blockEntry.username);
				if (!spiHelper_isCheckuser() && overrideExisting &&
					spiHelper_CU_BLOCK_RE.exec(blockReason)) {
					// If you're not a checkuser, we've asked to overwrite existing blocks, and the block
					// target has a CU block on them, check whether that was intended
					if (!confirm('User ' + blockEntry.username + ' appears to be CheckUser-blocked, are you SURE you want to re-block them?\n' +
						'Current block message:\n' + blockReason
					)) {
						return;
					}
				}
				const isIP = mw.util.isIPAddress(blockEntry.username, true);
				const isIPRange = isIP && !mw.util.isIPAddress(blockEntry.username, false);
				let blockSummary = 'Abusing [[WP:SOCK|multiple accounts]]: Please see: [[' + spiHelper_interwikiPrefix + spiHelper_pageName + ']]';
				if (spiHelper_isCheckuser() && cuBlock) {
					const cublock_template = isIP ? ('{{checkuserblock}}') : ('{{checkuserblock-account}}');
					if (cuBlockOnly) {
						blockSummary = cublock_template;
					} else {
						blockSummary = cublock_template + ': ' + blockSummary;
					}
				} else if (isIPRange) {
					blockSummary = '{{rangeblock| ' + blockSummary +
						(blockEntry.acb ? '' : '|create=yes') + '}}';
				}
				const blockSuccess = await spiHelper_blockUser(
					blockEntry.username,
					blockEntry.duration,
					blockSummary,
					overrideExisting,
					(isIP ? blockEntry.ab : false),
					blockEntry.acb,
					(isIP ? false : blockEntry.ab),
					blockEntry.ntp,
					blockEntry.nem,
					spiHelper_settings.watchBlockedUser,
					spiHelper_settings.watchBlockedUserExpiry);
				if (!blockSuccess) {
					// Don't add a block notice if we failed to block
					if (blockEntry.tpn) {
						// Also warn the user if we were going to post a block notice on their talk page
						const $statusLine = $('<li>').appendTo($('#spiHelper_status', document));
						$statusLine.addClass('spiHelper-errortext').html('<b>Block failed on ' + blockEntry.username + ', not adding talk page notice</b>');
					}
					return;
				}
				if (blockedList) {
					blockedList += ', ';
				}
				blockedList += '{{noping|' + blockEntry.username + '}}';
				
				if (isIPRange) {
					// There isn't really a talk page for an IP range, so return here before we reach that section
					return;
				}
				// Talk page notice
				if (blockEntry.tpn) {
					let newText = '';
					let isSock = blockEntry.tpn.includes('sock');
					// Hacky workaround for when we didn't make a master tag
					if (isSock && blockEntry.username === sockmaster) {
						isSock = false;
					}
					if (isSock) {
						newText = '== Blocked as a sockpuppet ==\n';
					} else {
						newText = '== Blocked for sockpuppetry ==\n';
					}
					newText += '{{subst:uw-sockblock|spi=' + spiHelper_caseName;
					if (blockEntry.duration === 'indefinite') {
						newText += '|indef=yes';
					} else {
						newText += '|duration=' + blockEntry.duration;
					}
					if (blockEntry.ntp) {
						newText += '|notalk=yes';
					}
					newText += '|sig=yes';
					if (isSock) {
						newText += '|master=' + sockmaster;
					}
					newText += '}}';

					if (!blankTalk) {
						const oldtext = await spiHelper_getPageText('User talk:' + blockEntry.username, true);
						if (oldtext !== '') {
							newText = oldtext + '\n' + newText;
						}
					}
					// Hardcode the watch setting to 'nochange' since we will have either watched or not watched based on the _boolean_
					// watchBlockedUser
					spiHelper_editPage('User talk:' + blockEntry.username,
						newText, 'Adding sockpuppetry block notice per [[' + spiHelper_getInterwikiPrefix() + spiHelper_pageName + ']]', false, 'nochange');
				}
			});
		}
		if (blockedList) {
			logMessage += '\n** blocked ' + blockedList;
		}

		let tagged = '';
		if (sockmaster) {
			// Whether we should purge sock pages (needed when we create a category)
			let needsPurge = false;
			// True for each we need to check if the respective category (e.g.
			// "Suspected sockpuppets of Test") exists
			let checkConfirmedCat = false;
			let checkSuspectedCat = false;
			let checkAltSuspectedCat = false;
			let checkAltConfirmedCat = false;
			spiHelper_tags.forEach(async (tagEntry) => {
				if (mw.util.isIPAddress(tagEntry.username, true)) {
					return; // do not support tagging IPs
				}
				let tagText = '';
				let altmasterName = '';
				let altmasterTag = '';
				if (altmaster !== '' && tagEntry.altmasterTag !== '') {
					altmasterName = altmaster;
					altmasterTag = tagEntry.altmasterTag;
					switch (altmasterTag) {
						case 'suspected':
							checkAltSuspectedCat = true;
							break;
						case 'proven':
							checkAltConfirmedCat = true;
							break;
					}
				}
				let isMaster = false;
				let tag = '';
				let checked = '';
				switch (tagEntry.tag) {
					case 'blocked':
						tag = 'blocked';
						checkSuspectedCat = true;
						break;
					case 'proven':
						tag = 'proven';
						checkConfirmedCat = true;
						break;
					case 'confirmed':
						tag = 'confirmed';
						checkConfirmedCat = true;
						break;
					case 'master':
						tag = 'blocked';
						isMaster = true;
						break;
					case 'sockmasterchecked':
						tag = 'blocked';
						checked = 'yes';
						isMaster = true;
						break;
					case 'bannedmaster':
						tag = 'banned';
						checked = 'yes';
						isMaster = true;
						break;
				}
				const isLocked = await spiHelper_isUserGloballyLocked(tagEntry.username) ? 'yes' : 'no';
				let isNotBlocked;
				// If this account is going to be blocked, force isNotBlocked to 'no' - it's possible that the
				// block hasn't gone through by the time we reach this point
				if (tagEntry.blocking) {
					isNotBlocked = 'no';
				} else {
					// Otherwise, query whether the user is blocked
					isNotBlocked = await spiHelper_getUserBlockReason(tagEntry.username) ? 'no' : 'yes';
				}
				if (isMaster) {
					// Not doing SPI or LTA fields for now - those auto-detect right now
					// and I'm not sure if setting them to empty would mess that up
					tagText += `{{sockpuppeteer
| 1 = ${tag}
| checked = ${checked}
| locked = ${isLocked}
}}`;
				}
				// Not if-else because we tag something as both sock and master if they're a
				// sockmaster and have a suspected altmaster
				if (!isMaster || altmasterName) {
					let sockmasterName = sockmaster;
					if (altmasterName && isMaster) {
						// If we have an altmaster and we're the master, swap a few values around
						sockmasterName = altmasterName;
						tag = altmasterTag;
						altmasterName = '';
						altmasterTag = '';
						tagText += '\n';
					}
					tagText += `{{sockpuppet
| 1 = ${sockmasterName}
| 2 = ${tag}
| locked = ${isLocked}
| notblocked = ${isNotBlocked}
| altmaster = ${altmasterName}
| altmaster-status = ${altmasterTag}
}}`;
				}
				spiHelper_editPage('User:' + tagEntry.username, tagText, 'Adding sockpuppetry tag per [[' + spiHelper_getInterwikiPrefix() + spiHelper_pageName + ']]',
					false, spiHelper_settings.watchTaggedUser, spiHelper_settings.watchTaggedUserExpiry);
				if (tagged) {
					tagged += ', ';
				}
				tagged += '{{noping|' + tagEntry.username + '}}';
			});
			if (tagged) {
				logMessage += '\n** tagged ' + tagged;
			}

			if (checkAltConfirmedCat) {
				const catname = 'Category:Wikipedia sockpuppets of ' + altmaster;
				const cattext = await spiHelper_getPageText(catname, false);
				// Empty text means the page doesn't exist - create it
				if (!cattext) {
					await spiHelper_editPage(catname, '{{sockpuppet category}}',
						'Creating sockpuppet category per [[' + spiHelper_getInterwikiPrefix() + spiHelper_pageName + ']]',
						true, spiHelper_settings.watchNewCats, spiHelper_settings.watchNewCatsExpiry);
					needsPurge = true;
				}
			}
			if (checkAltSuspectedCat) {
				const catname = 'Category:Suspected Wikipedia sockpuppets of ' + altmaster;
				const cattext = await spiHelper_getPageText(catname, false);
				if (!cattext) {
					await spiHelper_editPage(catname, '{{sockpuppet category}}',
						'Creating sockpuppet category per [[' + spiHelper_getInterwikiPrefix() + spiHelper_pageName + ']]',
						true, spiHelper_settings.watchNewCats, spiHelper_settings.watchNewCatsExpiry);
					needsPurge = true;
				}
			}
			if (checkConfirmedCat) {
				const catname = 'Category:Wikipedia sockpuppets of ' + sockmaster;
				const cattext = await spiHelper_getPageText(catname, false);
				if (!cattext) {
					await spiHelper_editPage(catname, '{{sockpuppet category}}',
						'Creating sockpuppet category per [[' + spiHelper_getInterwikiPrefix() + spiHelper_pageName + ']]',
						true, spiHelper_settings.watchNewCats, spiHelper_settings.watchNewCatsExpiry);
					needsPurge = true;
				}
			}
			if (checkSuspectedCat) {
				const catname = 'Category:Suspected Wikipedia sockpuppets of ' + sockmaster;
				const cattext = await spiHelper_getPageText(catname, false);
				if (!cattext) {
					await spiHelper_editPage(catname, '{{sockpuppet category}}',
						'Creating sockpuppet category per [[' + spiHelper_getInterwikiPrefix() + spiHelper_pageName + ']]',
						true, spiHelper_settings.watchNewCats, spiHelper_settings.watchNewCatsExpiry);
					needsPurge = true;
				}
			}
			// Purge the sock pages if we created a category (to get rid of
			// the issue where the page says "click here to create category"
			// when the category was created after the page)
			if (needsPurge) {
				spiHelper_tags.forEach((tagEntry) => {
					if (mw.util.isIPAddress(tagEntry.username, true)) {
						// Skip, this is an IP
						return;
					}
					if (!tagEntry.tag && !tagEntry.altmasterTag) {
						// Skip, not tagged
						return;
					}
					// Not bothering with an await, no need for async behavior here
					spiHelper_purgePage('User:' + tagEntry.username);
				});
			}
		}
		if (spiHelper_globalLocks.length > 0) {
			let locked = '';
			let templateContent = '';
			let matchCount = 0;
			spiHelper_globalLocks.forEach(async (globalLockEntry) => {
				// do not support locking IPs (those are global blocks, not
				// locks, and are handled a bit differently)
				if (mw.util.isIPAddress(globalLockEntry, true)) {
					return;
				}
				templateContent += '|' + globalLockEntry;
				if (locked) {
					locked += ', ';
				}
				locked += '{{noping|' + globalLockEntry + '}}';
				matchCount++;
			});

			if (matchCount > 0) {
				if (hideLockNames) {
					// If requested, hide locked names
					templateContent += '|hidename=1';
				}
				// Parts of this code were adapted from https://github.com/Xi-Plus/twinkle-global
				let lockTemplate = '';
				if (matchCount === 1) {
					lockTemplate = '{{LockHide' + templateContent + '}}';
				} else {
					lockTemplate = '{{MultiLock' + templateContent + '}}';
				}
				if (!sockmaster) {
					sockmaster = prompt('Please enter the name of the sockmaster: ', spiHelper_caseName);
				}
				const lockComment = prompt('Please enter a comment for the global lock request (optional):', '');
				const heading = hideLockNames ? 'sockpuppet(s)' : '[[Special:CentralAuth/' + sockmaster + '|' + sockmaster + ']] sock(s)';
				let message = '=== Global lock for ' + heading + ' ===';
				message += '\n{{status}}';
				message += '\n' + lockTemplate;
				message += '\n* Sockpuppet(s) found in enwiki sockpuppet investigation, see [[' + spiHelper_interwikiPrefix + spiHelper_pageName + ']]. ' + lockComment + ' ~~~~';

				// Write lock request to [[meta:Steward requests/Global]]
				let srgText = await spiHelper_getPageText('meta:Steward requests/Global', false);
				srgText = srgText.replace(/\n+(== See also == *\n)/, '\n\n' + message + '\n\n$1');
				spiHelper_editPage('meta:Steward requests/Global', srgText, 'global lock request for ' + heading, false, 'nochange');
				$statusAnchor.append($('<li>').text('Filing global lock request'));
			}
			if (locked) {
				logMessage += '\n** requested locks for ' + locked;
			}
		}
	}
	if (spiHelper_sectionId && comment && comment !== '*') {
		if (!sectionText.includes('\n----')) {
			sectionText += '\n----<!-- All comments go ABOVE this line, please. -->';
		}
		if (!/~~~~/.test(comment)) {
			comment += ' ~~~~';
		}
		// Clerks and admins post in the admin section
		if (spiHelper_isClerk() || spiHelper_isAdmin()) {
			// Complicated regex to find the first regex in the admin section
			// The weird (\n|.) is because we can't use /s (dot matches newline) regex mode without ES9,
			// I don't want to go there yet
			sectionText = sectionText.replace(/----(?!(\n|.)*----)/, comment + '\n----');
		} else { // Everyone else posts in the "other users" section
			sectionText = sectionText.replace(spiHelper_ADMIN_SECTION_RE,
				'\n' + comment + '\n====<big>Clerk, CheckUser, and/or patrolling admin comments</big>====\n');
		}
		if (editsummary) {
			editsummary += ', comment';
		} else {
			editsummary = 'Comment';
		}
		logMessage += '\n** commented';
	}

	if (spiHelper_ActionsSelected.Close) {
		newCaseStatus = 'close';
		if (editsummary) {
			editsummary += ', marking case as closed';
		} else {
			editsummary = 'Marking case as closed';
		}
		logMessage += '\n** closed case';
	}
	if (spiHelper_sectionId !== null) {
		const caseStatusText = spiHelper_CASESTATUS_RE.exec(sectionText)[0];
		sectionText = sectionText.replace(caseStatusText, '{{SPI case status|' + newCaseStatus + '}}');
	}

	// Fallback: if we somehow managed to not make an edit summary, add a default one
	if (!editsummary) {
		editsummary = 'Saving page';
	}

	// Make all of the requested edits (synchronous since we might make more changes to the page)
	await spiHelper_editPage(spiHelper_pageName, sectionText, editsummary, false,
		spiHelper_settings.watchCase, spiHelper_settings.watchCaseExpiry, spiHelper_startingRevID, spiHelper_sectionId);
	// Update to the latest revision ID
	spiHelper_startingRevID = await spiHelper_getPageRev(spiHelper_pageName);
	if (spiHelper_ActionsSelected.Archive) {
		// Archive the case
		if (spiHelper_sectionId === null) {
			// Archive the whole case
			logMessage += '\n** Archived case';
			await spiHelper_archiveCase();
		} else {
			// Just archive the selected section
			logMessage += '\n** Archived section';
			await spiHelper_archiveCaseSection(spiHelper_sectionId);
		}
	} else if (spiHelper_ActionsSelected.Rename && renameTarget) {
		if (spiHelper_sectionId === null) {
			// Option 1: we selected "All cases," this is a whole-case move/merge
			logMessage += '\n** moved/merged case to ' + renameTarget;
			await spiHelper_moveCase(renameTarget);
		} else {
			// Option 2: this is a single-section case move or merge
			logMessage += '\n** moved section to ' + renameTarget;
			await spiHelper_moveCaseSection(renameTarget, spiHelper_sectionId);
		}
	}
	if (spiHelper_settings.log) {
		spiHelper_log(logMessage);
	}

	await spiHelper_purgePage(spiHelper_pageName);
	$('#spiHelper_status', document).append($('<li>').text('Done!'));
	spiHelper_activeOperations.set('mainActions', 'successful');
}

/**
 * Logs SPI actions to userspace a la Twinkle's CSD/prod/etc. logs
 *
 * @param {string} logString String with the changes the user made
 */
async function spiHelper_log(logString) {
	const now = new Date();
	const dateString = now.toLocaleString('en', { month: 'long' }) + ' ' +
		now.toLocaleString('en', { year: 'numeric' });
	const dateHeader = '==\\s*' + dateString + '\\s*==';
	const dateHeaderRe = new RegExp(dateHeader, 'i');

	let logPageText = await spiHelper_getPageText('User:' + mw.config.get('wgUserName') + '/spihelper_log', false);
	if (!logPageText.match(dateHeaderRe)) {
		logPageText += '\n== ' + dateString + ' ==';
	}
	logPageText += '\n' + logString;
	await spiHelper_editPage('User:' + mw.config.get('wgUserName') + '/spihelper_log', logPageText, 'Logging spihelper edits', false, 'nochange');
}

// Major helper functions
/**
 * Cleanups following a rename - update the archive notice, add an archive notice to the
 * old case name, add the original sockmaster to the sock list for reference
 *
 * @param {string} oldCasePage Title of the previous case page
 */
async function spiHelper_postRenameCleanup(oldCasePage) {
	'use strict';
	const replacementArchiveNotice = '<noinclude>__TOC__</noinclude>\n{{SPIarchive notice|' + spiHelper_caseName + '}}\n{{SPIpriorcases}}';
	const oldCaseName = oldCasePage.replace(/Wikipedia:Sockpuppet investigations\//g, '');

	// The old case should just be the archivenotice template and point to the new case
	spiHelper_editPage(oldCasePage, replacementArchiveNotice, 'Updating case following page move', false, spiHelper_settings.watchCase, spiHelper_settings.watchCaseExpiry);

	// The new case's archivenotice should be updated with the new name
	let newPageText = await spiHelper_getPageText(spiHelper_pageName, true);
	newPageText = newPageText.replace(spiHelper_ARCHIVENOTICE_RE, '{{SPIarchive notice|' + spiHelper_caseName + '}}');
	// We also want to add the previous master to the sock list
	// We use SOCK_SECTION_RE_WITH_NEWLINE to clean up any extraneous whitespace
	newPageText = newPageText.replace(spiHelper_SOCK_SECTION_RE_WITH_NEWLINE, '====Suspected sockpuppets====' +
		'\n* {{checkuser|1=' + oldCaseName + '}} ({{clerknote}} original case name)\n');
	// Also remove the new master if they're in the sock list
	// This RE is kind of ugly. The idea is that we find everything from the level 4 heading
	// ending with "sockpuppets" to the level 4 heading beginning with <big> and pull the checkuser
	// template matching the current case name out. This keeps us from accidentally replacing a
	// checkuser entry in the admin section
	const newMasterReString = '(sockpuppets\\s*====.*?)\\n^\\s*\\*\\s*{{checkuser\\|(?:1=)?' + spiHelper_caseName + '(?:\\|master name\\s*=.*?)?}}\\s*$(.*====\\s*<big>)';
	const newMasterRe = new RegExp(newMasterReString, 'sm');
	newPageText = newPageText.replace(newMasterRe, '$1\n$2');

	await spiHelper_editPage(spiHelper_pageName, newPageText, 'Updating case following page move', false, spiHelper_settings.watchCase, spiHelper_settings.watchCaseExpiry, spiHelper_startingRevID);
	// Update to the latest revision ID
	spiHelper_startingRevID = await spiHelper_getPageRev(spiHelper_pageName);
}

/**
 * Cleanups following a merge - re-insert the original page text
 *
 * @param {string} originalText Text of the page pre-merge
 */
async function spiHelper_postMergeCleanup(originalText) {
	'use strict';
	let newText = await spiHelper_getPageText(spiHelper_pageName, false);
	// Remove the SPI header templates from the page
	newText = newText.replace(/\n*<noinclude>__TOC__.*\n/ig, '');
	newText = newText.replace(spiHelper_ARCHIVENOTICE_RE, '');
	newText = newText.replace(spiHelper_PRIORCASES_RE, '');
	newText = originalText + '\n' + newText;

	// Write the updated case
	await spiHelper_editPage(spiHelper_pageName, newText, 'Re-adding previous cases following merge', false, spiHelper_settings.watchCase, spiHelper_settings.watchCaseExpiry, spiHelper_startingRevID);
	// Update to the latest revision ID
	spiHelper_startingRevID = await spiHelper_getPageRev(spiHelper_pageName);
}

/**
 * Archive all closed sections of a case
 */
async function spiHelper_archiveCase() {
	'use strict';
	let i = 0;
	let previousRev = 0;
	while (i < spiHelper_caseSections.length) {
		const sectionId = spiHelper_caseSections[i].index;
		const sectionText = await spiHelper_getPageText(spiHelper_pageName, false,
			sectionId);

		const currentRev = await spiHelper_getPageRev(spiHelper_pageName);
		if (previousRev === currentRev && currentRev !== 0) {
			// Our previous archive hasn't gone through yet, wait a bit and retry
			await new Promise(resolve => setTimeout(resolve, 100));

			// Re-grab the case sections list since the page may have updated
			spiHelper_caseSections = await spiHelper_getInvestigationSectionIDs();
			continue;
		}
		previousRev = await spiHelper_getPageRev(spiHelper_pageName);
		i++;
		const result = spiHelper_CASESTATUS_RE.exec(sectionText);
		if (result === null) {
			// Bail out - can't find the case status template in this section
			continue;
		}
		if (spiHelper_CASESTATUS_CLOSED_RE.test(result[1])) {
			// A running concern with the SPI archives is whether they exceed the post-expand
			// include size. Calculate what percent of that size the archive will be if we
			// add the current page to it - if >1, we need to archive the archive
			const postExpandPercent =
				(await spiHelper_getPostExpandSize(spiHelper_pageName, sectionId) +
				await spiHelper_getPostExpandSize(spiHelper_getArchiveName())) /
				spiHelper_getMaxPostExpandSize();
			if (postExpandPercent >= 1) {
				// We'd overflow the archive, so move it and then archive the current page
				// Find the first empty archive page
				let archiveId = 1;
				while (await spiHelper_getPageText(spiHelper_getArchiveName() + '/' + archiveId, false) !== '') {
					archiveId++;
				}
				const newArchiveName = spiHelper_getArchiveName() + '/' + archiveId;
				await spiHelper_movePage(spiHelper_getArchiveName(), newArchiveName, 'Moving archive to avoid exceeding post expand size limit', false);
			}
			// Need an await here - if we have multiple sections archiving we don't want
			// to stomp on each other
			await spiHelper_archiveCaseSection(sectionId);
			// need to re-fetch caseSections since the section numbering probably just changed,
			// also reset our index
			i = 0;
			spiHelper_caseSections = await spiHelper_getInvestigationSectionIDs();
		}
	}
}

/**
 * Archive a specific section of a case
 *
 * @param {!number} sectionId The section number to archive
 */
async function spiHelper_archiveCaseSection(sectionId) {
	'use strict';
	let sectionText = await spiHelper_getPageText(spiHelper_pageName, true, sectionId);
	sectionText = sectionText.replace(spiHelper_CASESTATUS_RE, '');
	const newarchivetext = sectionText.substring(sectionText.search(spiHelper_SECTION_RE));

	// Update the archive
	let archivetext = await spiHelper_getPageText(spiHelper_getArchiveName(), true);
	if (!archivetext) {
		archivetext = '__' + 'TOC__\n{{SPIarchive notice|1=' + spiHelper_caseName + '}}\n{{SPIpriorcases}}';
	} else {
		archivetext = archivetext.replace(/<br\s*\/>\s*{{SPIpriorcases}}/gi, '\n{{SPIpriorcases}}'); // fmt fix whenever needed.
	}
	archivetext += '\n' + newarchivetext;
	const archiveSuccess = await spiHelper_editPage(spiHelper_getArchiveName(), archivetext,
		'Archiving case section from [[' + spiHelper_getInterwikiPrefix() + spiHelper_pageName + ']]',
		false, spiHelper_settings.watchArchive, spiHelper_settings.watchArchiveExpiry);

	if (!archiveSuccess) {
		const $statusLine = $('<li>').appendTo($('#spiHelper_status', document));
		$statusLine.addClass('spiHelper-errortext').html('<b>Failed to update archive, not removing section from case page</b>');
		return;
	}
		
	// Blank the section we archived
	await spiHelper_editPage(spiHelper_pageName, '', 'Archiving case section to [[' + spiHelper_getInterwikiPrefix() + spiHelper_getArchiveName() + ']]',
		false, spiHelper_settings.watchCase, spiHelper_settings.watchCaseExpiry, spiHelper_startingRevID, sectionId);
	// Update to the latest revision ID
	spiHelper_startingRevID = await spiHelper_getPageRev(spiHelper_pageName);
}

/**
 * Move or merge the selected case into a different case
 *
 * @param {string} target The username portion of the case this section should be merged into
 *                        (should have been normalized before getting passed in)
 */
async function spiHelper_moveCase(target) {
	// Move or merge an entire case
	// Normalize: change underscores to spaces
	target = target;
	const newPageName = spiHelper_pageName.replace(spiHelper_caseName, target);
	const targetPageText = await spiHelper_getPageText(newPageName, false);
	if (targetPageText) {
		if (spiHelper_isAdmin()) {
			const proceed = confirm('Target page exists, do you want to histmerge the cases?');
			if (!proceed) {
				// Build out the error line
				$('<li>')
					.append($('<div>').addClass('spihelper-errortext')
						.append($('<b>').text('Aborted merge.')))
					.appendTo($('#spiHelper_status', document));
				return;
			}
		} else {
			$('<li>')
				.append($('<div>').addClass('spihelper-errortext')
					.append($('<b>').text('Target page exists and you are not an admin, aborting merge.')))
				.appendTo($('#spiHelper_status', document));
			return;
		}
	}
	// Housekeeping to update all of the var names following the rename
	const oldPageName = spiHelper_pageName;
	const oldArchiveName = spiHelper_getArchiveName();
	spiHelper_caseName = target;
	spiHelper_pageName = newPageName;
	let archivesCopied = false;
	if (targetPageText) {
		// There's already a page there, we're going to merge
		// First, check if there's an archive; if so, copy its text over
		const newArchiveName = spiHelper_getArchiveName().replace(spiHelper_caseName, target);
		let sourceArchiveText = await spiHelper_getPageText(oldArchiveName, false);
		let targetArchiveText = await spiHelper_getPageText(newArchiveName, false);
		if (sourceArchiveText && targetArchiveText) {
			$('<li>')
			.append($('<div>').text('Archive detected on both source and target cases, manually copying archive.'))
			.appendTo($('#spiHelper_status', document));

			// Normalize the source archive text
			sourceArchiveText = sourceArchiveText.replace(/^\s*__TOC__\s*$\n/gm, '');
			sourceArchiveText = sourceArchiveText.replace(spiHelper_ARCHIVENOTICE_RE, '');
			sourceArchiveText = sourceArchiveText.replace(spiHelper_PRIORCASES_RE, '');
			// Strip leading newlines
			sourceArchiveText = sourceArchiveText.replace(/^\n*/, '');
			targetArchiveText += '\n' + sourceArchiveText;
			await spiHelper_editPage(newArchiveName, targetArchiveText, 'Copying archives from [[' + spiHelper_getInterwikiPrefix() + oldArchiveName + ']], see page history for attribution',
				false, spiHelper_settings.watchArchive, spiHelper_settings.watchArchiveExpiry);
			await spiHelper_deletePage(oldArchiveName, 'Deleting copied archive');
			archivesCopied = true;
		}
		// Ignore warnings on the move, we're going to get one since we're stomping an existing page
		await spiHelper_deletePage(spiHelper_pageName, 'Deleting as part of case merge');
		await spiHelper_movePage(oldPageName, spiHelper_pageName, 'Merging case to [[' + spiHelper_getInterwikiPrefix() + spiHelper_pageName + ']]', true);
		await spiHelper_undeletePage(spiHelper_pageName, 'Restoring page history after merge');
		if (archivesCopied) {
			// Create a redirect
			spiHelper_editPage(oldArchiveName, '#REDIRECT [[' + newArchiveName + ']]', 'Redirecting old archive to new archive',
				false, spiHelper_settings.watchArchive, spiHelper_settings.watchArchiveExpiry);
		}
	} else {
		await spiHelper_movePage(oldPageName, spiHelper_pageName, 'Moving case to [[' + spiHelper_getInterwikiPrefix() + spiHelper_pageName + ']]', false);
	}
	spiHelper_startingRevID = await spiHelper_getPageRev(spiHelper_pageName);
	await spiHelper_postRenameCleanup(oldPageName);
	if (targetPageText) {
		// If there was a page there before, also need to do post-merge cleanup
		await spiHelper_postMergeCleanup(targetPageText);
	}
	if (archivesCopied) {
		alert('Archives were merged during the case move, please reorder the archive sections');
	}
}

/**
 * Move or merge a specific section of a case into a different case
 *
 * @param {string} target The username portion of the case this section should be merged into (pre-normalized)
 * @param {!number} sectionId The section ID of this case that should be moved/merged
 */
async function spiHelper_moveCaseSection(target, sectionId) {
	// Move or merge a particular section of a case
	'use strict';
	const newPageName = spiHelper_pageName.replace(spiHelper_caseName, target);
	let targetPageText = await spiHelper_getPageText(newPageName, false);
	let sectionText = await spiHelper_getPageText(spiHelper_pageName, true, sectionId);
	// SOCK_SECTION_RE_WITH_NEWLINE cleans up extraneous whitespace at the top of the section
	// Have to do this transform before concatenating with targetPageText so that the
	// "originally filed" goes in the correct section
	sectionText = sectionText.replace(spiHelper_SOCK_SECTION_RE_WITH_NEWLINE, '====Suspected sockpuppets====' +
	'\n* {{checkuser|1=' + spiHelper_caseName + '}} ({{clerknote}} originally filed under this user)\n');

	if (targetPageText === '') {
		// Pre-load the split target with the SPI templates if it's empty
		targetPageText = '<noinclude>__TOC__</noinclude>\n{{SPIarchive notice|' + target + '}}\n{{SPIpriorcases}}';
	}
	targetPageText += '\n' + sectionText;

	// Intentionally not async - doesn't matter when this edit finishes
	spiHelper_editPage(newPageName, targetPageText, 'Moving case section from [[' + spiHelper_getInterwikiPrefix() + spiHelper_pageName + ']], see page history for attribution',
		false, spiHelper_settings.watchCase, spiHelper_settings.watchCaseExpiry);
	// Blank the section we moved
	await spiHelper_editPage(spiHelper_pageName, '', 'Moving case section to [[' + spiHelper_getInterwikiPrefix() + newPageName + ']]',
		false, spiHelper_settings.watchCase, spiHelper_settings.watchCaseExpiry, spiHelper_startingRevID, sectionId);
	// Update to the latest revision ID
	spiHelper_startingRevID = await spiHelper_getPageRev(spiHelper_pageName);
}

/**
 * Render a text box's contents and display it in the preview area
 *
 */
async function spiHelper_previewText() {
	const inputText = $('#spiHelper_CommentText', document).val().toString();
	const renderedText = await spiHelper_renderText(spiHelper_pageName, inputText);
	// Fill the preview box with the new text
	const $previewBox = $('#spiHelper_previewBox', document);
	$previewBox.html(renderedText);
	// Unhide it if it was hidden
	$previewBox.show();
}

/**
 * Given a page title, get an API to operate on that page
 *
 * @param {string} title Title of the page we want the API for
 * @return {Object} MediaWiki Api/ForeignAPI for the target page's wiki
 */
function spiHelper_getAPI(title) {
	'use strict';
	if (title.startsWith('m:') || title.startsWith('meta:')) {
		return new mw.ForeignApi('https://meta.wikimedia.org/w/api.php');
	} else {
		return new mw.Api();
	}
}

/**
 * Removes the interwiki prefix from a page title
 *
 * @param {*} title Page name including interwiki prefix
 * @return {string} Just the page name
 */
function spiHelper_stripXWikiPrefix(title) {
	// TODO: This only works with single-colon names, make it more robust
	'use strict';
	if (title.startsWith('m:') || title.startsWith('meta:')) {
		return title.slice(title.indexOf(':') + 1);
	} else {
		return title;
	}
}

/**
 * Get the post-expand include size of a given page
 *
 * @param {string} title Page title to check
 * @param {?number} sectionId Section to check, if null check the whole page
 *
 * @return {Promise<number>} Post-expand include size of the given page/page section
 */
async function spiHelper_getPostExpandSize(title, sectionId = null) {
	// Synchronous method to get a page's post-expand include size given its title
	const finalTitle = spiHelper_stripXWikiPrefix(title);

	const request = {
		action: 'parse',
		prop: 'limitreportdata',
		page: finalTitle
	};
	if (sectionId) {
		request.section = sectionId;
	}
	const api = spiHelper_getAPI(title);
	try {
		const response = await api.get(request);

		// The page might not exist, so we need to handle that smartly - only get the parse
		// if the page actually parsed
		if ('parse' in response) {
			// Iterate over all properties to find the PEIS
			for (let i = 0; i < response.parse.limitreportdata.length; i++) {
				if (response.parse.limitreportdata[i].name === 'limitreport-postexpandincludesize') {
					return response.parse.limitreportdata[i][0];
				}
			}
		} else {
			// Fallback - most likely the page doesn't exist
			return 0;
		}
	} catch (error) {
		// Something's gone wrong, just return 0
		return 0;

	}
}

/**
 * Get the maximum post-expand size from the wgPageParseReport (it's the same for all pages)
 *
 * @return {number} The max post-expand size in bytes
 */
function spiHelper_getMaxPostExpandSize() {
	'use strict';
	return mw.config.get('wgPageParseReport').limitreport.postexpandincludesize.limit;
}

/**
 * Get the inter-wiki prefix for the current wiki
 *
 * @return {string} The inter-wiki prefix
 */
function spiHelper_getInterwikiPrefix() {
	// Mostly copied from https://github.com/Xi-Plus/twinkle-global/blob/master/morebits.js
	// Most of this should be overkill (since most of these wikis don't have checkuser support)
	/** @type {string[]} */ const temp = mw.config.get('wgServer').replace(/^(https?)?\/\//, '').split('.');
	const wikiLang = temp[0];
	const wikiFamily = temp[1];
	switch (wikiFamily) {
		case 'wikimedia':
			switch (wikiLang) {
				case 'commons':
					return ':commons:';
				case 'meta':
					return ':meta:';
				case 'species:':
					return ':species:';
				case 'incubator':
					return ':incubator:';
				default:
					return '';
			}
		case 'mediawiki':
			return 'mw';
		case 'wikidata:':
			switch (wikiLang) {
				case 'test':
					return ':testwikidata:';
				case 'www':
					return ':d:';
				default:
					return '';
			}
		case 'wikipedia':
			switch (wikiLang) {
				case 'test':
					return ':testwiki:';
				case 'test2':
					return ':test2wiki:';
				default:
					return ':w:' + wikiLang + ':';
			}
		case 'wiktionary':
			return ':wikt:' + wikiLang + ':';
		case 'wikiquote':
			return ':q:' + wikiLang + ':';
		case 'wikibooks':
			return ':b:' + wikiLang + ':';
		case 'wikinews':
			return ':n:' + wikiLang + ':';
		case 'wikisource':
			return ':s:' + wikiLang + ':';
		case 'wikiversity':
			return ':v:' + wikiLang + ':';
		case 'wikivoyage':
			return ':voy:' + wikiLang + ':';
		default:
			return '';
	}
}

// "Building-block" functions to wrap basic API calls
/**
 * Get the text of a page. Not that complicated.
 *
 * @param {string} title Title of the page to get the contents of
 * @param {boolean} show Whether to show page fetch progress on-screen
 * @param {?number} [sectionId=null] Section to retrieve, setting this to null will retrieve the entire page
 *
 * @return {Promise<string>} The text of the page, '' if the page does not exist.
 */
async function spiHelper_getPageText(title, show, sectionId = null) {
	const $statusLine = $('<li>');
	if (show) {
		// Actually display the statusLine
		$('#spiHelper_status', document).append($statusLine);
	}
	// Build the link element (use JQuery so we get escapes and such)
	const $link = $('<a>').attr('href', mw.util.getUrl(title)).attr('title', title).text(title);
	$statusLine.html('Getting page ' + $link.prop('outerHTML'));

	const finalTitle = spiHelper_stripXWikiPrefix(title);

	const request = {
		action: 'query',
		prop: 'revisions',
		rvprop: 'content',
		rvslots: 'main',
		indexpageids: true,
		titles: finalTitle
	};

	if (sectionId) {
		request.rvsection = sectionId;
	}

	try {
		const response = await spiHelper_getAPI(title).get(request);
		const pageid = response.query.pageids[0];

		if (pageid === '-1') {
			$statusLine.html('Page ' + $link.html() + ' does not exist');
			return '';
		}
		$statusLine.html('Got ' + $link.html());
		return response.query.pages[pageid].revisions[0].slots.main['*'];
	} catch (error) {
		$statusLine.addClass('spiHelper-errortext').html('<b>Failed to get ' + $link.html() + '</b>: ' + error);
		return '';
	}
}

/**
 *
 * @param {string} title Title of the page to edit
 * @param {string} newtext New content of the page
 * @param {string} summary Edit summary to use for the edit
 * @param {boolean} createonly Only try to create the page - if false,
 *                             will fail if the page already exists
 * @param {string} watch What watchlist setting to use when editing - decides
 *                       whether the edited page will be watched
 * @param {string} watchExpiry Duration to watch the edited page, if unset
 *                             defaults to 'indefinite'
 * @param {?number} baseRevId Base revision ID, used to detect edit conflicts. If null,
 *                           we'll grab the current page ID.
 * @param {?number} [sectionId=null] Section to edit - if null, edits the whole page
 * 
 * @return {Promise<boolean>} Whether the edit was successful
 */
async function spiHelper_editPage(title, newtext, summary, createonly, watch, watchExpiry = null, baseRevId = null, sectionId = null) {
	let activeOpKey = 'edit_' + title;
	if (sectionId) {
		activeOpKey += '_' + sectionId;
	}
	spiHelper_activeOperations.set(activeOpKey, 'running');
	const $statusLine = $('<li>').appendTo($('#spiHelper_status', document));
	const $link = $('<a>').attr('href', mw.util.getUrl(title)).attr('title', title).text(title);

	$statusLine.html('Editing ' + $link.prop('outerHTML'));

	if (!baseRevId) {
		baseRevId = await spiHelper_getPageRev(title);
	}
	const api = spiHelper_getAPI(title);
	const finalTitle = spiHelper_stripXWikiPrefix(title);

	const request = {
		action: 'edit',
		watchlist: watch,
		summary: summary + spihelper_ADVERT,
		text: newtext,
		title: finalTitle,
		createonly: createonly,
		baserevid: baseRevId
	};
	if (sectionId) {
		request.section = sectionId;
	}
	if (watchExpiry) {
		request.watchlistExpiry = watchExpiry;
	}
	try {
		await api.postWithToken('csrf', request);
		$statusLine.html('Saved ' + $link.prop('outerHTML'));
		spiHelper_activeOperations.set(activeOpKey, 'success');
		return true;
	} catch (error) {
		$statusLine.addClass('spiHelper-errortext').html('<b>Edit failed on ' + $link.html() + '</b>: ' + error);
		console.error(error);
		spiHelper_activeOperations.set(activeOpKey, 'failed');
		return false;
	}
}
/**
 * Moves a page. Exactly what it sounds like.
 *
 * @param {string} sourcePage Title of the source page (page we're moving)
 * @param {string} destPage Title of the destination page (page we're moving to)
 * @param {string} summary Edit summary to use for the move
 * @param {boolean} ignoreWarnings Whether to ignore warnings on move (used to force-move one page over another)
 */
async function spiHelper_movePage(sourcePage, destPage, summary, ignoreWarnings) {
	// Move a page from sourcePage to destPage. Not that complicated.
	'use strict';

	let activeOpKey = 'move_' + sourcePage + '_' + destPage;
	spiHelper_activeOperations.set(activeOpKey, 'running');

	// Should never be a crosswiki call
	const api = new mw.Api();

	const $statusLine = $('<li>').appendTo($('#spiHelper_status', document));
	const $sourceLink = $('<a>').attr('href', mw.util.getUrl(sourcePage)).attr('title', sourcePage).text(sourcePage);
	const $destLink = $('<a>').attr('href', mw.util.getUrl(destPage)).attr('title', destPage).text(destPage);

	$statusLine.html('Moving ' + $sourceLink.prop('outerHTML') + ' to ' + $destLink.prop('outerHTML'));

	try {
		await api.postWithToken('csrf', {
			action: 'move',
			from: sourcePage,
			to: destPage,
			reason: summary + spihelper_ADVERT,
			noredirect: false,
			movesubpages: true,
			ignoreWarnings: ignoreWarnings
		});
		$statusLine.html('Moved ' + $sourceLink.prop('outerHTML') + ' to ' + $destLink.prop('outerHTML'));
		spiHelper_activeOperations.set(activeOpKey, 'success');
	} catch (error) {
		$statusLine.addClass('spihelper-errortext').html('<b>Failed to move ' + $sourceLink.prop('outerHTML') + ' to ' + $destLink.prop('outerHTML') + '</b>: ' + error);
		spiHelper_activeOperations.set(activeOpKey, 'failed');
	}
}

/**
 * Purges a page's cache
 *
 *
 * @param {string} title Title of the page to purge
 */
async function spiHelper_purgePage(title) {
	// Forces a cache purge on the selected page
	'use strict';
	const $statusLine = $('<li>').appendTo($('#spiHelper_status', document));
	const $link = $('<a>').attr('href', mw.util.getUrl(title)).attr('title', title).text(title);
	$statusLine.html('Purging ' + $link.prop('outerHTML'));
	const strippedTitle = spiHelper_stripXWikiPrefix(title);

	const api = spiHelper_getAPI(title);
	try {
		await api.postWithToken('csrf', {
			action: 'purge',
			titles: strippedTitle
		});
		$statusLine.html('Purged ' + $link.prop('outerHTML'));
	} catch (error) {
		$statusLine.addClass('spihelper-errortext').html('<b>Failed to purge ' + $link.prop('outerHTML') + '</b>: ' + error);
	}
}

/**
 * Blocks a user.
 *
 * @param {string} user Username to block
 * @param {string} duration Duration of the block
 * @param {string} reason Reason to log for the block
 * @param {boolean} reblock Whether to reblock - if false, nothing will happen if the target user is already blocked
 * @param {boolean} anononly For IPs, whether this is an anonymous-only block (alternative is
 *                           that logged-in users with the IP are also blocked)
 * @param {boolean} accountcreation Whether to permit the user to create new accounts
 * @param {boolean} autoblock Whether to apply an autoblock to the user's IP
 * @param {boolean} talkpage Whether to revoke talkpage access
 * @param {boolean} email Whether to block email
 * @param {boolean} watchBlockedUser Watchlist setting for whether to watch the newly-blocked user
 * @param {string} watchExpiry Duration to watch the blocked user, if unset
 *                             defaults to 'indefinite'

 * @return {Promise<boolean>} True if the block suceeded, false if not
 */
async function spiHelper_blockUser(user, duration, reason, reblock, anononly, accountcreation,
	autoblock, talkpage, email, watchBlockedUser, watchExpiry) {
	'use strict';
	let activeOpKey = 'block_' + user;
	spiHelper_activeOperations.set(activeOpKey, 'running');

	if (!watchExpiry) {
		watchExpiry = 'indefinite';
	}
	const userPage = 'User:' + user;
	const $statusLine = $('<li>').appendTo($('#spiHelper_status', document));
	const $link = $('<a>').attr('href', mw.util.getUrl(userPage)).attr('title', userPage).text(user);
	$statusLine.html('Blocking ' + $link.prop('outerHTML'));

	// This is not something which should ever be cross-wiki
	const api = new mw.Api();
	try {
		await api.postWithToken('csrf', {
			action: 'block',
			expiry: duration,
			reason: reason,
			reblock: reblock,
			anononly: anononly,
			nocreate: accountcreation,
			autoblock: autoblock,
			allowusertalk: !talkpage,
			noemail: email,
			watchuser: watchBlockedUser,
			watchlistexpiry: watchExpiry,
			user: user
		});
		$statusLine.html('Blocked ' + $link.prop('outerHTML'));
		spiHelper_activeOperations.set(activeOpKey, 'success');
		return true;
	} catch (error) {
		$statusLine.addClass('spihelper-errortext').html('<b>Failed to block ' + $link.prop('outerHTML') + '</b>: ' + error);
		spiHelper_activeOperations.set(activeOpKey, 'failed');
		return false;
	}
}

/**
 * Get whether a user is currently blocked
 *
 * @param {string} user Username
 * @return {Promise<string>} Block reason, empty string if not blocked
 */
async function spiHelper_getUserBlockReason(user) {
	'use strict';
	// This is not something which should ever be cross-wiki
	const api = new mw.Api();
	try {
		const response = await api.get({
			action: 'query',
			list: 'blocks',
			bklimit: '1',
			bkusers: user,
			bkprop: 'user|reason'
		});
		if (response.query.blocks.length === 0) {
			// If the length is 0, then the user isn't blocked
			return '';
		}
		return response.query.blocks[0].reason;
	} catch (error) {
		return '';
	}
}

/**
 * Get whether a user is currently globally locked
 *
 * @param {string} user Username
 * @return {Promise<boolean>} Whether the user is globally locked
 */
async function spiHelper_isUserGloballyLocked(user) {
	'use strict';
	// This is not something which should ever be cross-wiki
	const api = new mw.Api();
	try {
		const response = await api.get({
			action: 'query',
			list: 'globalallusers',
			agulimit: '1',
			agufrom: user,
			aguto: user,
			aguprop: 'lockinfo'
		});
		if (response.query.globalallusers.length === 0) {
			// If the length is 0, then we couldn't find the global user
			return false;
		}
		// If the 'locked' field is present, then the user is locked
		return 'locked' in response.query.globalallusers[0];
	} catch (error) {
		return false;
	}
}

/**
 * Get a page's latest revision ID - useful for preventing edit conflicts
 *
 * @param {string} title Title of the page
 * @return {Promise<number>} Latest revision of a page, 0 if it doesn't exist
 */
async function spiHelper_getPageRev(title) {
	'use strict';

	const finalTitle = spiHelper_stripXWikiPrefix(title);
	const request = {
		action: 'query',
		prop: 'revisions',
		rvslots: 'main',
		indexpageids: true,
		titles: finalTitle
	};

	try {
		const response = await spiHelper_getAPI(title).get(request);
		const pageid = response.query.pageids[0];
		if (pageid === '-1') {
			return 0;
		}
		return response.query.pages[pageid].revisions[0].revid;
	} catch (error) {
		return 0;
	}
}

/**
 * Delete a page. Admin-only function.
 *
 * @param {string} title Title of the page to delete
 * @param {string} reason Reason to log for the page deletion
 */
async function spiHelper_deletePage(title, reason) {
	'use strict';

	let activeOpKey = 'delete_' + title;
	spiHelper_activeOperations.set(activeOpKey, 'running');

	const $statusLine = $('<li>').appendTo($('#spiHelper_status', document));
	const $link = $('<a>').attr('href', mw.util.getUrl(title)).attr('title', title).text(title);
	$statusLine.html('Deleting ' + $link.prop('outerHTML'));

	const api = spiHelper_getAPI(title);
	try {
		await api.postWithToken('csrf', {
			action: 'delete',
			title: title,
			reason: reason
		});
		$statusLine.html('Deleted ' + $link.prop('outerHTML'));
		spiHelper_activeOperations.set(activeOpKey, 'success');
	} catch (error) {
		$statusLine.addClass('spihelper-errortext').html('<b>Failed to delete ' + $link.prop('outerHTML') + '</b>: ' + error);
		spiHelper_activeOperations.set(activeOpKey, 'failed');
	}
}

/**
 * Undelete a page (or, if the page exists, undelete deleted revisions). Admin-only function
 *
 * @param {string} title Title of the pgae to undelete
 * @param {string} reason Reason to log for the page undeletion
 */
async function spiHelper_undeletePage(title, reason) {
	'use strict';
	let activeOpKey = 'undelete_' + title;
	spiHelper_activeOperations.set(activeOpKey, 'running');

	const $statusLine = $('<li>').appendTo($('#spiHelper_status', document));
	const $link = $('<a>').attr('href', mw.util.getUrl(title)).attr('title', title).text(title);
	$statusLine.html('Undeleting ' + $link.prop('outerHTML'));

	const api = spiHelper_getAPI(title);
	try {
		await api.postWithToken('csrf', {
			action: 'undelete',
			title: title,
			reason: reason
		});
		$statusLine.html('Undeleted ' + $link.prop('outerHTML'));
		spiHelper_activeOperations.set(activeOpKey, 'success');
	} catch (error) {
		$statusLine.addClass('spihelper-errortext').html('<b>Failed to undelete ' + $link.prop('outerHTML') + '</b>: ' + error);
		spiHelper_activeOperations.set(activeOpKey, 'failed');
	}
}

/**
 * Render a snippet of wikitext
 *
 * @param {string} title Page title
 * @param {string} text Text to render
 * @return {Promise<string>} Rendered version of the text
 */
async function spiHelper_renderText(title, text) {
	'use strict';

	const request = {
		action: 'parse',
		prop: 'text',
		pst: 'true',
		text: text,
		title: title
	};

	try {
		const response = await spiHelper_getAPI(title).get(request);
		return response.parse.text['*'];
	} catch (error) {
		console.error('Error rendering text: ' + error);
		return '';
	}
}

/**
 * Get a list of investigations on the sockpuppet investigation page
 *
 * @return {Promise<Object[]>} An array of section objects, each section is a separate investigation
 */
async function spiHelper_getInvestigationSectionIDs() {
	// Uses the parse API to get page sections, then find the investigation
	// sections (should all be level-3 headers)
	'use strict';

	// Since this only affects the local page, no need to call spiHelper_getAPI()
	const api = new mw.Api();
	const response = await api.get({
		action: 'parse',
		prop: 'sections',
		page: spiHelper_pageName
	});
	const dateSections = [];
	for (let i = 0; i < response.parse.sections.length; i++) {
		// TODO: also check for presence of spi case status
		if (response.parse.sections[i].level === '3') {
			dateSections.push(response.parse.sections[i]);
		}
	}
	return dateSections;
}

/**
 * Pretty obvious - gets the name of the archive. This keeps us from having to regen it
 * if we rename the case
 *
 * @return {string} Name of the archive page
 */
function spiHelper_getArchiveName() {
	return spiHelper_pageName + '/Archive';
}

// UI helper functions
/**
 * Generate a line of the block table for a particular user
 *
 * @param {string} name Username for this block line
 * @param {boolean} defaultblock Whether to check the block box by default on this row
 * @param {number} id Index of this line in the block table
 */
function spiHelper_generateBlockTableLine(name, defaultblock, id) {
	'use strict';

	const $table = $('#spiHelper_blockTable', document);

	const $row = $('<tr>');
	// Username
	$('<td>').append($('<input>').attr('type', 'text').attr('id', 'spiHelper_block_username' + id)
		.val(name).addClass('.spihelper-widthlimit')).appendTo($row);
	// Block checkbox (only for admins)
	$('<td>').addClass('spiHelper_adminClass').append($('<input>').attr('type', 'checkbox')
		.attr('id', 'spiHelper_block_doblock' + id).prop('checked', defaultblock)).appendTo($row);
	// Block duration (only for admins)
	const defaultBlockDuration = mw.util.isIPAddress(name, true) ? '1 week' : 'indefinite';
	$('<td>').addClass('spiHelper_adminClass').append($('<input>').attr('type', 'text')
		.attr('id', 'spiHelper_block_duration' + id).val(defaultBlockDuration)
		.addClass('.spihelper-widthlimit')).appendTo($row);
	// Account creation blocked (only for admins)
	$('<td>').addClass('spiHelper_adminClass').append($('<input>').attr('type', 'checkbox')
		.attr('id', 'spiHelper_block_acb' + id).prop('checked', true)).appendTo($row);
	// Autoblock (only for admins)
	$('<td>').addClass('spiHelper_adminClass').append($('<input>').attr('type', 'checkbox')
		.attr('id', 'spiHelper_block_ab' + id).prop('checked', true)).appendTo($row);
	// Revoke talk page access (only for admins)
	$('<td>').addClass('spiHelper_adminClass').append($('<input>').attr('type', 'checkbox')
		.attr('id', 'spiHelper_block_tp' + id)).appendTo($row);
	// Block email access (only for admins)
	$('<td>').addClass('spiHelper_adminClass').append($('<input>').attr('type', 'checkbox')
		.attr('id', 'spiHelper_block_email' + id)).appendTo($row);
	// Tag select box
	$('<td>').append($('<select>').attr('id', 'spiHelper_block_tag' + id)
		.val(name)).appendTo($row);
	// Altmaster tag select
	$('<td>').append($('<select>').attr('id', 'spiHelper_block_tag_altmaster' + id)
		.val(name)).appendTo($row);
	// Global lock (disabled for IPs since they can't be locked)
	$('<td>').append($('<input>').attr('type', 'checkbox').attr('id', 'spiHelper_block_lock' + id)
		.prop('disabled', mw.util.isIPAddress(name, true))).appendTo($row);
	$table.append($row);

	// Generate the select entries
	spiHelper_generateSelect('spiHelper_block_tag' + id, spiHelper_TAG_OPTIONS);
	spiHelper_generateSelect('spiHelper_block_tag_altmaster' + id, spiHelper_ALTMASTER_TAG_OPTIONS);
}

/**
 * Complicated function to decide what checkboxes to enable or disable
 * and which to check by default
 */
async function spiHelper_setCheckboxesBySection() {
	// Displays the top-level SPI menu
	'use strict';

	const $topView = $('#spiHelper_topViewDiv', document);
	// Get the value of the selection box
	if ($('#spiHelper_sectionSelect', $topView).val() === 'all') {
		spiHelper_sectionId = null;
		spiHelper_sectionName = null;
	} else {
		spiHelper_sectionId = parseInt($('#spiHelper_sectionSelect', $topView).val().toString());
		const $sectionSelect = $('#spiHelper_sectionSelect', $topView);
		spiHelper_sectionName = spiHelper_caseSections[$sectionSelect.prop('selectedIndex')].line;
	}

	const $archiveBox = $('#spiHelper_Archive', $topView);
	const $blockBox = $('#spiHelper_BlockTag', $topView);
	const $closeBox = $('#spiHelper_Close', $topView);
	const $commentBox = $('#spiHelper_Comment', $topView);
	const $moveBox = $('#spiHelper_Move', $topView);
	const $caseActionBox = $('#spiHelper_Case_Action', $topView);
	const $spiMgmtBox = $('#spiHelper_SpiMgmt', $topView);


	// Start by unchecking everything
	$archiveBox.prop('checked', false);
	$blockBox.prop('checked', false);
	$closeBox.prop('checked', false);
	$commentBox.prop('checked', false);
	$moveBox.prop('checked', false);
	$caseActionBox.prop('checked', false);
	$spiMgmtBox.prop('checked', false);

	// Enable optionally-disabled boxes
	$closeBox.prop('disabled', false);
	$archiveBox.prop('disabled', false);

	if (spiHelper_sectionId === null) {
		// Hide inputs that aren't relevant in the case view
		$('.spiHelper_singleCaseOnly', $topView).hide();
		// Show inputs only visible in all-case mode
		$('.spiHelper_allCasesOnly', $topView).show();
		// Fix the move label
		$('#spiHelper_moveLabel', $topView).text('Move/merge full case (Clerk only)');
		// enable the move box
		$moveBox.prop('disabled', false);
	} else {
		const sectionText = await spiHelper_getPageText(spiHelper_pageName, false, spiHelper_sectionId);
		if (!spiHelper_SECTION_RE.test(sectionText)) {
			// Nothing to do here.
			return;
		}
		
		// Unhide single-case options
		$('.spiHelper_singleCaseOnly', $topView).show();
		// Hide inputs only visible in all-case mode
		$('.spiHelper_allCasesOnly', $topView).hide();

		const result = spiHelper_CASESTATUS_RE.exec(sectionText);
		let casestatus = '';
		if (result) {
			casestatus = result[1];
		}

		// Disable the section move setting if you haven't opted into it
		if (!spiHelper_settings.iUnderstandSectionMoves) {
			$moveBox.prop('disabled', true);
		}

		const isClosed = spiHelper_CASESTATUS_CLOSED_RE.test(casestatus);

		if (isClosed) {
			$closeBox.prop('disabled', true);
			$archiveBox.prop('checked', true);
		} else {
			$archiveBox.prop('disabled', true);
		}

		// Change the label on the rename button
		$('#spiHelper_moveLabel', $topView).html('Move case section (<span title="You probably want to move the full case, ' +
			'select All Sections instead of a specific date in the drop-down"' +
			'class="rt-commentedText spihelper-hovertext"><b>READ ME FIRST</b></span>)');
	}
}

/**
 * Updates whether the 'archive' checkbox is enabled
 */
function spiHelper_updateArchive() {
	// Archive should only be an option if close is checked or disabled (disabled meaning that
	// the case is closed) and rename is not checked
	'use strict';
	$('#spiHelper_Archive', document).prop('disabled', !($('#spiHelper_Close', document).prop('checked') ||
		$('#spiHelper_Close', document).prop('disabled')) || $('#spiHelper_Move', document).prop('checked'));
	if ($('#spiHelper_Archive', document).prop('disabled')) {
		$('#spiHelper_Archive', document).prop('checked', false);
	}
}

/**
 * Updates whether the 'move' checkbox is enabled
 */
function spiHelper_updateMove() {
	// Rename is mutually exclusive with archive
	'use strict';
	$('#spiHelper_Move', document).prop('disabled', $('#spiHelper_Archive', document).prop('checked'));
	if ($('#spiHelper_Move', document).prop('disabled')) {
		$('#spiHelper_Move', document).prop('checked', false);
	}
}

/**
 * Generate a select input, optionally with an onChange call
 *
 * @param {string} id Name of the input
 * @param {SelectOption[]} options Array of options objects
 */
function spiHelper_generateSelect(id, options) {
	// Add the dates to the selector
	const $selector = $('#' + id, document);
	for (let i = 0; i < options.length; i++) {
		const o = options[i];
		$('<option>')
			.val(o.value)
			.prop('selected', o.selected)
			.text(o.label)
			.prop('disabled', o.disabled)
			.appendTo($selector);
	}
}

/**
 * Given an HTML element, sets that element's value on all block options
 * For example, checking the 'block all' button will check all per-user 'block' elements
 *
 * @param {JQuery<HTMLElement>} source The HTML input element that we're matching all selections to
 */
function spiHelper_setAllBlockOpts(source) {
	'use strict';
	for (let i = 1; i <= spiHelper_usercount; i++) {
		const target = $('#' + source.attr('id') + i);
		if (source.attr('type') === 'checkbox') {
			// Don't try to set disabled checkboxes
			if (!target.prop('disabled')) {
				target.prop('checked', source.prop('checked'));
			}
		} else {
			target.val(source.val());
		}
	}
}

/**
 * Inserts text at the cursor's position
 *
 * @param {JQuery<HTMLElement>} source Select box that was changed
 * @param {number?} pos Position to insert text; if null, inserts at the cursor
 */
function spiHelper_insertTextFromSelect(source, pos = null) {
	const $textBox = $('#spiHelper_CommentText', document);
	// https://stackoverflow.com/questions/11076975/how-to-insert-text-into-the-textarea-at-the-current-cursor-position
	const selectionStart = parseInt($textBox.attr('selectionStart'));
	const selectionEnd = parseInt($textBox.attr('selectionEnd'));
	const startText = $textBox.val().toString();
	const newText = source.val().toString();
	if (pos === null && (selectionStart || selectionStart === 0)) {
		$textBox.val(startText.substring(0, selectionStart) +
			newText +
			startText.substring(selectionEnd, startText.length));
		$textBox.attr('selectionStart', selectionStart + newText.length);
		$textBox.attr('selectionEnd', selectionEnd + newText.length);
	} else if (pos !== null) {
		$textBox.val(startText.substring(0, pos) +
			source.val() +
			startText.substring(pos, startText.length));
		$textBox.attr('selectionStart', selectionStart + newText.length);
		$textBox.attr('selectionEnd', selectionEnd + newText.length);
	} else {
		$textBox.val(startText + newText);
	}

	// Force the selected element to reset its selection to 0
	source.prop('selectedIndex', 0);
}

/**
 * Inserts a {{note}} template at the start of the text box
 *
 * @param {JQuery<HTMLElement>} source Select box that was changed
 */
function spiHelper_insertNote(source) {
	'use strict';
	const $textBox = $('#spiHelper_CommentText', document);
	let newText = $textBox.val().toString();
	// Match the start of the line, optionally including a '*' with or without whitespace around it,
	// optionally including a template which contains the string "note"
	newText = newText.replace(/^(\s*\*\s*)?({{[\w\s]*note[\w\s]*}}\s*)?/i, '* ' + '{{' + source.val() + '}} ');
	$textBox.val(newText);

	// Force the selected element to reset its selection to 0
	source.prop('selectedIndex', 0);
}

/**
 * Changes the case status in the comment box
 *
 * @param {JQuery<HTMLElement>} source Select box that was changed
 */
function spiHelper_caseActionUpdated(source) {
	const $textBox = $('#spiHelper_CommentText', document);
	const oldText = $textBox.val().toString();
	let newTemplate = '';
	switch (source.val()) {
		case 'CUrequest':
			newTemplate = '{{CURequest}}';
			break;
		case 'admin':
			newTemplate = '{{awaitingadmin}}';
			break;
		case 'clerk':
			newTemplate = '{{Clerk Request}}';
			break;
		case 'selfendorse':
			newTemplate = '{{Requestandendorse}}';
			break;
		case 'inprogress':
			newTemplate = '{{Inprogress}}';
			break;
		case 'decline':
			newTemplate = '{{Decline}}';
			break;
		case 'cudecline':
			newTemplate = '{{Cudecline}}';
			break;
		case 'endorse':
			newTemplate = '{{Endorse}}';
			break;
		case 'cuendorse':
			newTemplate = '{{cuendorse}}';
			break;
		case 'moreinfo': // Intentional fallthrough
		case 'cumoreinfo':
			newTemplate = '{{moreinfo}}';
			break;
		case 'relist':
			newTemplate = '{{relisted}}';
			break;
		case 'hold':
		case 'cuhold':
			newTemplate = '{{onhold}}';
			break;
	}
	if (spiHelper_CLERKSTATUS_RE.test(oldText)) {
		$textBox.val(oldText.replace(spiHelper_CLERKSTATUS_RE, newTemplate));
		if (!newTemplate) { // If the new template is empty, get rid of the stray ' - '
			$textBox.val(oldText.replace(/^ - /, ''));
		}
	} else if (newTemplate) {
		// Don't try to insert if the "new template" is empty
		// Also remove the leading *
		$textBox.val('*' + newTemplate + ' - ' + oldText.replace(/^\s*\*\s*/, ''));
	}
}

/**
 * Fires on page load, adds the SPI portlet and (if the page is categorized as "awaiting
 * archive," meaning that at least one closed template is on the page) the SPI-Archive portlet
 */
async function spiHelper_addLink() {
	'use strict';
	await spiHelper_loadSettings();
	await mw.loader.load('mediawiki.util');
	const initLink = mw.util.addPortletLink('p-cactions', '#', 'SPI', 'ca-spiHelper');
	initLink.addEventListener('click', (e) => {
		e.preventDefault();
		return spiHelper_init();
	});
	if (mw.config.get('wgCategories').includes('SPI cases awaiting archive') && spiHelper_isClerk()) {
		const oneClickArchiveLink = mw.util.addPortletLink('p-cactions', '#', 'SPI-Archive', 'ca-spiHelperArchive');
		oneClickArchiveLink.addEventListener('click', (e) => {
			e.preventDefault();
			return spiHelper_oneClickArchive();
		});
	}
	window.addEventListener('beforeunload', (e) => {
		const $actionView = $('#spiHelper_actionViewDiv', document);
		if ($actionView.length > 0) {
			e.preventDefault();
			return true;
		}

		// Make sure no operations are still in flight
		let isDirty = false;
		spiHelper_activeOperations.forEach((value, _0, _1) => {
			if (value === 'running') {
				isDirty = true;
			}
		});
		if (isDirty) {
			e.preventDefault();
			return true;
		}
	});
}

/**
 * Checks for the existence of Special:MyPage/spihelper-options.js, and if it exists,
 * loads the settings from that page.
 */
async function spiHelper_loadSettings() {
	// Dynamically load a user's settings
	// Borrowed from code I wrote for [[User:Headbomb/unreliable.js]]
	try {
		await mw.loader.getScript('/w/index.php?title=Special:MyPage/spihelper-options.js&action=raw&ctype=text/javascript');
		if (typeof spiHelperCustomOpts !== 'undefined') {
			Object.entries(spiHelperCustomOpts).forEach(([ k, v ]) => {
				spiHelper_settings[k] = v;
			});
		}
	} catch (error) {
		mw.log.error('Error retrieving your spihelper-options.js');
		// More detailed error in the console
		console.error('Error getting local spihelper-options.js: ' + error);
	}
}

// User role helper functions
/**
 * Whether the current user has admin permissions, used to determine
 * whether to show block options
 *
 * @return {boolean} Whether the current user is an admin
 */
function spiHelper_isAdmin() {
	if (spiHelper_settings.debugForceAdminState !== null) {
		return spiHelper_settings.debugForceAdminState;
	}
	return mw.config.get('wgUserGroups').includes('sysop');
}

/**
 * Whether the current user has checkuser permissions, used to determine
 * whether to show checkuser options
 *
 * @return {boolean} Whether the current user is a checkuser
 */

function spiHelper_isCheckuser() {
	if (spiHelper_settings.debugForceCheckuserState !== null) {
		return spiHelper_settings.debugForceCheckuserState;
	}
	return mw.config.get('wgUserGroups').includes('checkuser');
}

/**
 * Whether the current user is a clerk, used to determine whether to show
 * clerk options
 *
 * @return {boolean} Whether the current user is a clerk
 */
function spiHelper_isClerk() {
	// Assumption: checkusers should see clerk options. Please don't prove this wrong.
	return spiHelper_settings.clerk || spiHelper_isCheckuser();
}

/**
 * Common username normalization function
 * @param {string} username Username to normalize
 *
 * @return {string} Normalized username
 */
function spiHelper_normalizeUsername(username) {
	// Replace underscores with spaces
	username = username.replace('_', ' ');
	// Get rid of bad hidden characters
	username = username.replace(spiHelper_HIDDEN_CHAR_NORM_RE, '');
	// Remove leading and trailing spaces
	username = username.trim();
	if (mw.util.isIPAddress(username, true)) {
		// For IP addresses, capitalize them (really only applies to IPv6)
		username = username.toUpperCase();
	} else {
		// For actual usernames, make sure the first letter is capitalized
		username = username.charAt(0).toUpperCase() + username.slice(1);
	}
	return username;
}
// </nowiki>

/**
 * Parse key features from an archivenotice
 * @param {string} page Page to parse
 * 
 * @return {Promise<ParsedArchiveNotice>} Parsed archivenotice
 */
async function spiHelper_parseArchiveNotice(page) {
	const pagetext = await spiHelper_getPageText(page, false);
	const match = spiHelper_ARCHIVENOTICE_RE.exec(pagetext);
	const username = match[1];
	let deny = false;
	let xwiki = false;
	if (match[2]) {
		for (const entry of match[2].split('|')) {
			if (!entry) {
				// split in such a way that it's just a pipe
				continue;
			}
			const splitEntry = entry.split('=');
			if (splitEntry.length !== 2) {
				console.error('Malformed archivenotice parameter ' + entry);
				continue;
			}
			const key = splitEntry[0];
			const val = splitEntry[1];
			if (val.toLowerCase() !== 'yes') {
				// Only care if the value is 'yes'
				continue;
			}
			if (key.toLowerCase() === 'deny') {
				deny = true;
			} else if (key.toLowerCase() === 'crosswiki') {
				xwiki = true;
			}
		}
	}
	/** @type {ParsedArchiveNotice} */
	return {
		username: username,
		deny: deny,
		xwiki: xwiki
	};
}

/**
 * Helper function to make a new archivenotice
 * @param {string} username Username
 * @param {ParsedArchiveNotice} archiveNoticeParams Other archivenotice params
 * 
 * @return {string} New archivenotice
 */
 function spiHelper_makeNewArchiveNotice(username, archiveNoticeParams) {
	let notice = '{{SPIarchive notice|1=' + username;
	if (archiveNoticeParams.xwiki) {
		notice += '|crosswiki=yes';
	}
	if (archiveNoticeParams.deny) {
		notice += '|deny=yes';
	}
	notice += '}}';

	return notice;
 }
