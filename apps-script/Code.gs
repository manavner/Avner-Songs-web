/**
 * Avner Man Songs – sheet updater (Google Apps Script, bound to the songs sheet).
 *
 * Deployed as a Web App, it is called by the UPDATE button on the site:
 *   GET <web-app-url>?action=update
 * It looks for public YouTube videos by Avner Man (his own channel, the DistroKid
 * "Topic" channel, and a YouTube search by name) and appends the ones that are not
 * yet in the sheet to the bottom, continuing the numbering.
 *
 * Sheet columns (first tab): מספר | שם הסרטון | שם הערוץ | קישור YouTube
 */

// Channels whose videos belong in the list. Add more here if needed.
const CHANNELS = [
  { id: 'UC0b65NhvnAnsmIJaBnj1J1g', name: 'אבנר מן' },           // own uploads
  { id: 'UCz1luN4ETMqoUnBW6ojE42Q', name: 'Avner Man - Topic' }, // DistroKid releases
];

// YouTube search, to catch a new DistroKid "Topic" channel not listed above.
// Only results from a channel named exactly "<artist> - Topic" are added; other channels
// with similar names (e.g. "AvnerMania") are someone else's.
const SEARCH_QUERIES = ['Avner Man', 'אבנר מן'];
const CHANNEL_NAME_PATTERN = /^(avner man|אבנר מן) - topic$/i;

// Videos whose links appear anywhere in this tab are never added – put non-song uploads
// here when removing them from the list, so UPDATE doesn't add them back.
const IGNORE_SHEET_NAME = 'לא להוסיף';

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'update';
  let result;
  try {
    if (action !== 'update') throw new Error('Unknown action: ' + action);
    result = updateSongs();
  } catch (err) {
    result = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Run this once from the editor to authorize the script and test it. */
function updateSongs() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets()[0];
    const values = sheet.getDataRange().getValues();
    const header = values[0].map(String);
    const col = (re, fallback) => {
      const i = header.findIndex(h => re.test(h));
      return i >= 0 ? i : fallback;
    };
    const numCol = col(/^\s*(מספר|#|no\.?)\s*$/i, 0);
    const titleCol = col(/שם הסרטון|שם השיר|title/i, 1);
    const channelCol = col(/ערוץ|channel/i, 2);
    const linkCol = col(/קישור|link|url/i, 3);

    const existing = new Set();
    let lastNum = 0;
    values.slice(1).forEach(row => {
      const id = videoId(String(row[linkCol]));
      if (id) existing.add(id);
      const n = Number(row[numCol]);
      if (n > lastNum) lastNum = n;
    });

    const ignored = new Set();
    const ignoreSheet = ss.getSheetByName(IGNORE_SHEET_NAME);
    if (ignoreSheet) {
      ignoreSheet.getDataRange().getValues().forEach(row => row.forEach(cell => {
        const id = videoId(String(cell));
        if (id) ignored.add(id);
      }));
    }

    const found = typeof YouTube !== 'undefined' ? findWithApi() : findWithRss();
    const fresh = found.videos
      .filter(v => !existing.has(v.id) && !ignored.has(v.id))
      .sort((a, b) => String(a.publishedAt).localeCompare(String(b.publishedAt)));

    const width = Math.max(header.length, numCol + 1, titleCol + 1, channelCol + 1, linkCol + 1);
    const rows = fresh.map(v => {
      const row = new Array(width).fill('');
      row[numCol] = ++lastNum;
      row[titleCol] = v.title;
      row[channelCol] = v.channel;
      row[linkCol] = 'https://www.youtube.com/watch?v=' + v.id;
      return row;
    });
    if (rows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, width).setValues(rows);
    }

    return {
      ok: true,
      source: found.source,
      added: fresh.map(v => ({ title: v.title, channel: v.channel, url: 'https://www.youtube.com/watch?v=' + v.id })),
      total: existing.size + fresh.length,
    };
  } finally {
    lock.releaseLock();
  }
}

// ─── YouTube Data API (advanced service "YouTube") – finds all videos ───
function findWithApi() {
  const byId = new Map();
  const add = v => { if (v.id && !byId.has(v.id)) byId.set(v.id, v); };

  CHANNELS.forEach(ch => {
    try {
      // Every channel's uploads playlist id is its channel id with "UC" replaced by "UU".
      let pageToken;
      do {
        const res = YouTube.PlaylistItems.list('snippet,status', {
          playlistId: 'UU' + ch.id.slice(2), maxResults: 50, pageToken,
        });
        (res.items || []).forEach(it => {
          if (!it.status || it.status.privacyStatus !== 'public') return; // skip private/unlisted
          add({
            id: it.snippet.resourceId.videoId,
            title: it.snippet.title,
            channel: it.snippet.channelTitle || ch.name,
            publishedAt: it.snippet.publishedAt,
          });
        });
        pageToken = res.nextPageToken;
      } while (pageToken);
    } catch (err) {
      // Some Topic channels have no uploads playlist – fall back to a search by channel.
      searchVideos({ channelId: ch.id }).forEach(add);
    }
  });

  SEARCH_QUERIES.forEach(q => {
    searchVideos({ q }).filter(v => CHANNEL_NAME_PATTERN.test(v.channel)).forEach(add);
  });

  return { source: 'api', videos: [...byId.values()] };
}

function searchVideos(params) {
  const out = [];
  let pageToken, pages = 0;
  do {
    const res = YouTube.Search.list('snippet', Object.assign({ type: 'video', maxResults: 50, pageToken }, params));
    (res.items || []).forEach(it => out.push({
      id: it.id.videoId,
      title: decodeEntities(it.snippet.title), // search titles come HTML-escaped
      channel: decodeEntities(it.snippet.channelTitle),
      publishedAt: it.snippet.publishedAt,
    }));
    pageToken = res.nextPageToken;
  } while (pageToken && ++pages < 4); // each search page costs 100 of the 10,000 daily quota units
  return out;
}

// ─── Fallback without the YouTube service: RSS feeds (latest 15 videos per channel) ───
function findWithRss() {
  const videos = [];
  CHANNELS.forEach(ch => {
    const res = UrlFetchApp.fetch('https://www.youtube.com/feeds/videos.xml?channel_id=' + ch.id, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return;
    const root = XmlService.parse(res.getContentText()).getRootElement();
    const atom = XmlService.getNamespace('http://www.w3.org/2005/Atom');
    const yt = XmlService.getNamespace('yt', 'http://www.youtube.com/xml/schemas/2015');
    root.getChildren('entry', atom).forEach(entry => videos.push({
      id: entry.getChildText('videoId', yt),
      title: entry.getChildText('title', atom),
      channel: entry.getChild('author', atom).getChildText('name', atom) || ch.name,
      publishedAt: entry.getChildText('published', atom),
    }));
  });
  return { source: 'rss', videos };
}

function videoId(url) {
  const m = url.match(/(?:youtu\.be\/|youtube\.com\/(?:watch\?.*v=|embed\/|v\/|shorts\/))([a-zA-Z0-9_-]{11})/);
  return m ? m[1] : null;
}

function decodeEntities(s) {
  return String(s || '')
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>').replace(/&amp;/g, '&');
}
