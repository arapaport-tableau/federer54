// Photos supplied by the app owner (packaged under img/). Each entry:
//   key     caption type: serve | slice | forehand | usopen | backhand | pressure | trophy
//   hero    true = the opening photo (not a band)
//   after   section the band follows: hero | arc | margin | rivals | surf | shots | exp
//   focus   CSS object-position; side 'right' puts the caption on the right; layout 'split' for portrait photos
//   credit  photographer / licence line (required for Creative Commons images)
window.FED_PHOTOS = [
  { key: 'hero',     hero: true,      file: 'img/serve2.jpg',        focus: '38% 30%', alt: 'Federer in mid-air after striking a serve', credit: '' },
  { key: 'slice',    after: 'arc',    file: 'img/backhandslice.jpg', focus: '40% 8%', side: 'right', cap: 'top', alt: 'Federer preparing a backhand, Indian Wells', credit: '' },
  { key: 'forehand', after: 'margin', file: 'img/forehand1.jpg',     focus: '50% 8%', cap: 'top', alt: 'Federer setting up a forehand on grass at Wimbledon', credit: '' },
  { key: 'usopen',   after: 'rivals', file: 'img/volley1.jpg',       focus: '35% 22%', side: 'right', alt: 'Federer under the lights at the US Open', credit: '' },
  { key: 'backhand', after: 'surf',   file: 'img/backhand1.jpg',     focus: '50% 55%', side: 'right', alt: 'Federer finishing a one-handed backhand, US Open', credit: '' },
  { key: 'pressure', after: 'evo',    file: 'img/serve1.jpg',        layout: 'split', focus: '50% 55%', alt: 'Federer at the top of his service motion, US Open', credit: '' },
  { key: 'trophy',   after: 'exp',    file: 'img/winner1.jpg',       focus: '50% 10%', alt: 'Federer raising his arms to the crowd after a win, Rotterdam', credit: '' }
];
