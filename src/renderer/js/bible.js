// AnchorCast — Bible Database Engine
// Loads complete KJV (and other translations) from JSON files in /data/
// Falls back to built-in 223 key verses if JSON not yet installed.
// JSON format: [{"b":1,"c":1,"v":1,"t":"verse text"}, ...]  (b=book 1-66)
'use strict';

const CANON = {
  OT:['Genesis','Exodus','Leviticus','Numbers','Deuteronomy','Joshua','Judges','Ruth',
      '1 Samuel','2 Samuel','1 Kings','2 Kings','1 Chronicles','2 Chronicles',
      'Ezra','Nehemiah','Esther','Job','Psalms','Proverbs','Ecclesiastes',
      'Song of Solomon','Isaiah','Jeremiah','Lamentations','Ezekiel','Daniel',
      'Hosea','Joel','Amos','Obadiah','Jonah','Micah','Nahum','Habakkuk',
      'Zephaniah','Haggai','Zechariah','Malachi'],
  NT:['Matthew','Mark','Luke','John','Acts','Romans',
      '1 Corinthians','2 Corinthians','Galatians','Ephesians','Philippians',
      'Colossians','1 Thessalonians','2 Thessalonians','1 Timothy','2 Timothy',
      'Titus','Philemon','Hebrews','James','1 Peter','2 Peter',
      '1 John','2 John','3 John','Jude','Revelation']
};
const ALL_BOOKS = [...CANON.OT, ...CANON.NT];

const ABBREV = {
  gen:0,ex:1,exo:1,lev:2,num:3,deu:4,deut:4,josh:5,jdg:6,judg:6,rut:7,ruth:7,
  '1sa':8,'1sam':8,'2sa':9,'2sam':9,'1ki':10,'1kgs':10,'2ki':11,'2kgs':11,
  '1ch':12,'1chr':12,'2ch':13,'2chr':13,ezr:14,ezra:14,neh:15,est:16,esth:16,
  job:17,ps:18,psa:18,psalm:18,psalms:18,pss:18,pro:19,prov:19,ecc:20,eccl:20,
  song:21,sos:21,isa:22,jer:23,lam:24,eze:25,ezek:25,dan:26,hos:27,joe:28,joel:28,
  amo:29,amos:29,oba:30,obad:30,jon:31,jonah:31,mic:32,nah:33,hab:34,
  zep:35,zeph:35,hag:36,zec:37,zech:37,mal:38,
  mat:39,matt:39,mt:39,mar:40,mrk:40,mk:40,luk:41,lk:41,
  joh:42,jn:42,john:42,act:43,acts:43,rom:44,
  '1co':45,'1cor':45,'2co':46,'2cor':46,gal:47,eph:48,
  phi:49,phil:49,col:50,'1th':51,'1thes':51,'2th':52,'2thes':52,
  '1ti':53,'1tim':53,'2ti':54,'2tim':54,tit:55,phm:56,phlm:56,
  heb:57,jas:58,jam:58,'1pe':59,'1pet':59,'2pe':60,'2pet':60,
  '1jo':61,'1jn':61,'2jo':62,'2jn':62,'3jo':63,'3jn':63,jud:64,jude:64,rev:65
};

// Authoritative verse counts per chapter for all 66 books
const VERSE_COUNTS = [
  [31,25,24,26,32,22,24,22,29,32,32,20,18,24,21,16,27,33,38,18,34,24,20,67,34,35,46,22,35,43,55,32,20,31,29,43,36,30,23,23,57,38,34,34,28,34,31,22,33,26],
  [22,25,22,31,23,30,25,32,35,29,10,51,22,31,27,36,16,27,25,26,36,31,33,18,40,37,21,43,46,38,18,35,23,35,35,38,29,31,43,38],
  [17,16,17,35,19,30,38,36,24,20,47,8,59,57,33,34,16,30,24,16,15,18,21,32,32,33,30],
  [54,34,51,49,31,27,89,26,23,36,35,16,33,45,41,50,13,32,22,29,35,41,30,25,18,65,23,31,40,16,54,42,56,29,34,13],
  [46,37,29,49,33,25,26,20,29,22,32,32,18,29,23,22,20,22,21,20,23,30,25,22,19,19,26,68,29,20,30,52,29,12],
  [18,24,17,24,15,27,26,35,27,43,23,24,33,15,63,10,18,28,51,9,45,34,16,33],
  [36,23,31,24,31,40,25,35,57,18,40,15,25,20,20,31,13,31,30,48,25],
  [22,23,18,22],
  [28,36,21,22,12,21,17,22,27,27,15,25,23,52,35,23,58,30,24,42,15,23,29,22,44,25,12,25,11,31,13],
  [27,32,39,12,25,23,29,18,13,19,27,31,39,33,37,23,29,33,43,26,22,51,39,25],
  [53,46,28,34,18,38,51,66,28,29,43,33,34,31,34,34,24,46,21,43,29,53],
  [18,25,27,44,27,33,20,29,37,36,21,21,25,29,38,20,41,37,37,21,26,20,37,20,30],
  [54,55,24,43,26,81,40,40,44,14,47,40,14,17,29,43,27,17,19,8,30,19,32,31,31,32,34,21,30],
  [17,18,17,22,14,42,22,18,31,19,23,16,22,15,19,14,19,34,11,37,20,12,21,27,28,23,9,27,36,27,21,33,25,33,27,23],
  [11,70,13,24,17,22,28,36,15,44],
  [11,20,32,23,19,19,73,18,38,39,36,47,31],
  [22,23,15,17,14,14,10,17,32,3],
  [22,13,26,21,27,30,21,22,35,22,20,25,28,22,35,22,16,21,29,29,34,30,17,25,6,14,23,28,25,31,40,22,33,37,16,33,24,41,30,24,34,17],
  [6,12,8,8,12,10,17,9,20,18,7,8,6,7,5,11,15,50,14,9,13,31,6,10,22,12,14,9,11,12,24,11,22,22,28,12,40,22,13,17,13,11,5,20,28,22,35,22,20,32,22,20,24,11,22,33,22,28,33,2,20,43,1,45,6,11,44,29,9,44,37,9,47,31,13,25,9,37,21,34,8,47,32,13,11,1,16,14,23,31,8,10,24,20,28,34,26,6,12,15,18,21,10,14,21,16,17,22,16,15,28,18,14,22,33,21,17,24,18,33,16,13,13,11,1,3,8,14,51,17,16,15,14,26,16,14,9,32,24,9,35,21,23,14,18,14,9,13,24,12,29,18],
  [33,22,35,27,23,35,27,36,18,32,31,28,25,35,33,33,28,24,29,30,31,29,35,34,28,28,27,28,62,32,44],
  [18,26,22,16,20,12,29,17,18,20,10,14],
  [17,17,11,16,16,13,13,14],
  [31,22,26,6,30,13,25,22,21,34,16,6,22,32,9,14,14,7,25,6,17,25,18,23,12,21,13,29,24,33,9,20,24,17,10,22,38,22,8,31,29,25,28,28,25,13,15,22,26,11,23,15,12,17,13,12,21,14,21,22,11,12,19,12,25,24],
  [19,37,25,31,31,30,34,22,26,25,23,17,27,22,21,21,27,23,15,18,14,30,40,10,38,24,22,17,32,24,40,44,26,22,19,32,21,28,18,16,18,22,13,30,5,28,7,47,39,46,64,34],
  [22,22,66,22,22],
  [28,10,27,17,17,14,27,18,11,22,25,28,23,23,8,63,24,32,14,49,32,31,49,27,17,21,36,26,21,26,18,32,33,31,15,38,28,23,29,49,26,20,27,31,25,24,23,35],
  [21,49,30,37,31,28,28,27,27,21,45,13],
  [11,23,5,19,15,11,16,14,17,15,12,14,16,9],
  [20,32,21],[15,16,15,13,27,14,17,14,15],[21],[17,10,10,11],
  [16,13,12,13,15,16,20],[15,13,19],[17,20,19],[18,15,20],[15,23],
  [21,13,10,14,11,15,14,23,17,12,17,14,9,21],[14,17,18,6],
  [25,23,17,25,48,34,29,34,38,42,45,27,31,32,29,23,26,17,51,26,37,45,31,54,38,38,24,36],
  [45,28,35,41,43,56,37,38,50,52,33,44,37,72,47,20],
  [80,52,38,44,39,49,50,56,62,42,54,59,35,35,32,31,37,43,48,47,38,71,56,53],
  [51,25,36,54,47,71,53,59,41,42,57,50,38,31,27,33,26,40,42,31,25],
  [26,47,26,37,42,15,60,40,43,48,30,25,52,28,41,40,34,28,41,38,40,30,35,27,27,32,44,31],
  [32,29,31,25,21,23,25,39,33,21,36,21,14,26,33,24],
  [31,16,23,21,13,20,40,13,27,33,34,31,13,40,58,24],
  [24,17,18,18,21,18,16,24,15,18,33,21,14],
  [24,21,29,31,26,18],[23,22,21,28,30,14],[30,30,21,23],[29,23,25,18],
  [10,20,13,18,28],[12,17,18],[20,15,16,16,25,21],[18,26,17,22],[16,15,15],[25],
  [14,18,19,16,14,20,28,13,28,39,40,29,25],
  [27,26,18,17,20],[25,25,22,19,14],[21,22,18],[10,29,24,21,21],[13],[14],[25],
  [20,29,22,11,14,17,17,13,21,11,19,17,18,20,8,21,18,24,21,15,27,21]
];

// ─── IN-MEMORY STORES ────────────────────────────────────────────────────────
// Map key: "bookIdx:chapter:verse" (0-indexed book) → text string
// Dynamic translation store — any translation ID is supported
const STORES = {
  KJV:  new Map(),
  NKJV: new Map(),
};
// Live list of loaded translations — updated as data loads
let TRANSLATIONS = ['KJV','NKJV'];
let _dbLoaded = false;
let _dbStatus = 'fallback';

// Ensure a store exists for a given translation ID
function _ensureStore(id) {
  if (!STORES[id]) {
    STORES[id] = new Map();
    if (!TRANSLATIONS.includes(id)) TRANSLATIONS.push(id);
  }
}

// ─── BUILT-IN FALLBACK (223 key verses) ──────────────────────────────────────
// These are always available even without the JSON data file.
// Book index is 0-based to match ABBREV map.
function _seed() {
  const K = STORES.KJV, N = STORES.NKJV;
  const s = (b,c,v,k,n) => { K.set(`${b}:${c}:${v}`,k); if(n) N.set(`${b}:${c}:${v}`,n); };
  // GENESIS
  s(0,1,1,"In the beginning God created the heaven and the earth.","In the beginning God created the heavens and the earth.");
  s(0,1,27,"So God created man in his own image, in the image of God created he him; male and female created he them.","So God created man in His own image; in the image of God He created him; male and female He created them.");
  s(0,2,24,"Therefore shall a man leave his father and his mother, and shall cleave unto his wife: and they shall be one flesh.","Therefore a man shall leave his father and mother and be joined to his wife, and they shall become one flesh.");
  s(0,50,20,"But as for you, ye thought evil against me; but God meant it unto good.","But as for you, you meant evil against me; but God meant it for good.");
  // EXODUS
  s(1,14,14,"The LORD shall fight for you, and ye shall hold your peace.","The LORD will fight for you, and you shall hold your peace.");
  s(1,20,3,"Thou shalt have no other gods before me.","You shall have no other gods before Me.");
  // DEUTERONOMY
  s(4,6,5,"And thou shalt love the LORD thy God with all thine heart, and with all thy soul, and with all thy might.","You shall love the LORD your God with all your heart, with all your soul, and with all your strength.");
  s(4,31,6,"Be strong and of a good courage, fear not, nor be afraid of them: for the LORD thy God, he it is that doth go with thee; he will not fail thee, nor forsake thee.","Be strong and of good courage, do not fear nor be afraid of them; for the LORD your God, He is the One who goes with you. He will not leave you nor forsake you.");
  // JOSHUA
  s(5,1,9,"Have not I commanded thee? Be strong and of a good courage; be not afraid, neither be thou dismayed: for the LORD thy God is with thee whithersoever thou goest.","Have I not commanded you? Be strong and of good courage; do not be afraid, nor be dismayed, for the LORD your God is with you wherever you go.");
  s(5,24,15,"And if it seem evil unto you to serve the LORD, choose you this day whom ye will serve; but as for me and my house, we will serve the LORD.","And if it seems evil to you to serve the LORD, choose for yourselves this day whom you will serve; but as for me and my house, we will serve the LORD.");
  // 2 CHRONICLES
  s(13,7,14,"If my people, which are called by my name, shall humble themselves, and pray, and seek my face, and turn from their wicked ways; then will I hear from heaven, and will forgive their sin, and will heal their land.","if My people who are called by My name will humble themselves, and pray and seek My face, and turn from their wicked ways, then I will hear from heaven, and will forgive their sin and heal their land.");
  // JOB
  s(17,19,25,"For I know that my redeemer liveth, and that he shall stand at the latter day upon the earth:","For I know that my Redeemer lives, and He shall stand at last on the earth;");
  // PSALMS
  s(18,23,1,"The LORD is my shepherd; I shall not want.","The LORD is my shepherd; I shall not want.");
  s(18,23,2,"He maketh me to lie down in green pastures: he leadeth me beside the still waters.","He makes me to lie down in green pastures; He leads me beside the still waters.");
  s(18,23,3,"He restoreth my soul: he leadeth me in the paths of righteousness for his name's sake.","He restores my soul; He leads me in the paths of righteousness for His name's sake.");
  s(18,23,4,"Yea, though I walk through the valley of the shadow of death, I will fear no evil: for thou art with me; thy rod and thy staff they comfort me.","Yea, though I walk through the valley of the shadow of death, I will fear no evil; for You are with me; Your rod and Your staff, they comfort me.");
  s(18,23,5,"Thou preparest a table before me in the presence of mine enemies: thou anointest my head with oil; my cup runneth over.","You prepare a table before me in the presence of my enemies; You anoint my head with oil; my cup runs over.");
  s(18,23,6,"Surely goodness and mercy shall follow me all the days of my life: and I will dwell in the house of the LORD for ever.","Surely goodness and mercy shall follow me all the days of my life; and I will dwell in the house of the LORD forever.");
  s(18,27,1,"The LORD is my light and my salvation; whom shall I fear? the LORD is the strength of my life; of whom shall I be afraid?","The LORD is my light and my salvation; whom shall I fear? The LORD is the strength of my life; of whom shall I be afraid?");
  s(18,34,8,"O taste and see that the LORD is good: blessed is the man that trusteth in him.","Oh, taste and see that the LORD is good; blessed is the man who trusts in Him!");
  s(18,37,4,"Delight thyself also in the LORD; and he shall give thee the desires of thine heart.","Delight yourself also in the LORD, and He shall give you the desires of your heart.");
  s(18,46,1,"God is our refuge and strength, a very present help in trouble.","God is our refuge and strength, a very present help in trouble.");
  s(18,46,10,"Be still, and know that I am God: I will be exalted among the heathen, I will be exalted in the earth.","Be still, and know that I am God; I will be exalted among the nations, I will be exalted in the earth!");
  s(18,51,10,"Create in me a clean heart, O God; and renew a right spirit within me.","Create in me a clean heart, O God, and renew a steadfast spirit within me.");
  s(18,91,1,"He that dwelleth in the secret place of the most High shall abide under the shadow of the Almighty.","He who dwells in the secret place of the Most High shall abide under the shadow of the Almighty.");
  s(18,91,2,"I will say of the LORD, He is my refuge and my fortress: my God; in him will I trust.","I will say of the LORD, \"He is my refuge and my fortress; my God, in Him I will trust.\"");
  s(18,103,1,"Bless the LORD, O my soul: and all that is within me, bless his holy name.","Bless the LORD, O my soul; and all that is within me, bless His holy name!");
  s(18,118,24,"This is the day which the LORD hath made; we will rejoice and be glad in it.","This is the day the LORD has made; we will rejoice and be glad in it.");
  s(18,119,105,"Thy word is a lamp unto my feet, and a light unto my path.","Your word is a lamp to my feet and a light to my path.");
  s(18,121,1,"I will lift up mine eyes unto the hills, from whence cometh my help.","I will lift up my eyes to the hills— from whence comes my help?");
  s(18,121,2,"My help cometh from the LORD, which made heaven and earth.","My help comes from the LORD, who made heaven and earth.");
  s(18,139,14,"I will praise thee; for I am fearfully and wonderfully made: marvellous are thy works; and that my soul knoweth right well.","I will praise You, for I am fearfully and wonderfully made; marvelous are Your works, and that my soul knows very well.");
  // PROVERBS
  s(19,3,5,"Trust in the LORD with all thine heart; and lean not unto thine own understanding.","Trust in the LORD with all your heart, and lean not on your own understanding;");
  s(19,3,6,"In all thy ways acknowledge him, and he shall direct thy paths.","In all your ways acknowledge Him, and He shall direct your paths.");
  s(19,4,7,"Wisdom is the principal thing; therefore get wisdom: and with all thy getting get understanding.","Wisdom is the principal thing; therefore get wisdom. And in all your getting, get understanding.");
  s(19,18,21,"Death and life are in the power of the tongue: and they that love it shall eat the fruit thereof.","Death and life are in the power of the tongue, and those who love it will eat its fruit.");
  s(19,22,6,"Train up a child in the way he should go: and when he is old, he will not depart from it.","Train up a child in the way he should go, and when he is old he will not depart from it.");
  // ISAIAH
  s(22,40,28,"Hast thou not known? hast thou not heard, that the everlasting God, the LORD, the Creator of the ends of the earth, fainteth not, neither is weary?","Have you not known? Have you not heard? The everlasting God, the LORD, the Creator of the ends of the earth, neither faints nor is weary.");
  s(22,40,31,"But they that wait upon the LORD shall renew their strength; they shall mount up with wings as eagles; they shall run, and not be weary; and they shall walk, and not faint.","But those who wait on the LORD shall renew their strength; they shall mount up with wings like eagles, they shall run and not be weary, they shall walk and not faint.");
  s(22,41,10,"Fear thou not; for I am with thee: be not dismayed; for I am thy God: I will strengthen thee; yea, I will help thee; yea, I will uphold thee with the right hand of my righteousness.","Fear not, for I am with you; be not dismayed, for I am your God. I will strengthen you, yes, I will help you, I will uphold you with My righteous right hand.");
  s(22,53,5,"But he was wounded for our transgressions, he was bruised for our iniquities: the chastisement of our peace was upon him; and with his stripes we are healed.","But He was wounded for our transgressions, He was bruised for our iniquities; the chastisement for our peace was upon Him, and by His stripes we are healed.");
  s(22,55,11,"So shall my word be that goeth forth out of my mouth: it shall not return unto me void, but it shall accomplish that which I please.","So shall My word be that goes forth from My mouth; it shall not return to Me void, but it shall accomplish what I please, and it shall prosper in the thing for which I sent it.");
  // JEREMIAH
  s(23,29,11,"For I know the thoughts that I think toward you, saith the LORD, thoughts of peace, and not of evil, to give you an expected end.","For I know the thoughts that I think toward you, says the LORD, thoughts of peace and not of evil, to give you a future and a hope.");
  s(23,33,3,"Call unto me, and I will answer thee, and show thee great and mighty things, which thou knowest not.","Call to Me, and I will answer you, and show you great and mighty things, which you do not know.");
  // LAMENTATIONS
  s(24,3,22,"It is of the LORD's mercies that we are not consumed, because his compassions fail not.","Through the LORD's mercies we are not consumed, because His compassions fail not.");
  s(24,3,23,"They are new every morning: great is thy faithfulness.","They are new every morning; great is Your faithfulness.");
  // EZEKIEL
  s(25,36,26,"A new heart also will I give you, and a new spirit will I put within you.","I will give you a new heart and put a new spirit within you; I will take the heart of stone out of your flesh and give you a heart of flesh.");
  // MALACHI
  s(38,3,10,"Bring ye all the tithes into the storehouse, that there may be meat in mine house, and prove me now herewith, saith the LORD of hosts, if I will not open you the windows of heaven, and pour you out a blessing.","Bring all the tithes into the storehouse, that there may be food in My house, and try Me now in this, says the LORD of hosts, if I will not open for you the windows of heaven and pour out for you such blessing that there will not be room enough to receive it.");
  // MATTHEW
  s(39,5,3,"Blessed are the poor in spirit: for theirs is the kingdom of heaven.","Blessed are the poor in spirit, for theirs is the kingdom of heaven.");
  s(39,5,14,"Ye are the light of the world. A city that is set on an hill cannot be hid.","You are the light of the world. A city that is set on a hill cannot be hidden.");
  s(39,6,33,"But seek ye first the kingdom of God, and his righteousness; and all these things shall be added unto you.","But seek first the kingdom of God and His righteousness, and all these things shall be added to you.");
  s(39,11,28,"Come unto me, all ye that labour and are heavy laden, and I will give you rest.","Come to Me, all you who labor and are heavy laden, and I will give you rest.");
  s(39,16,18,"And I say also unto thee, That thou art Peter, and upon this rock I will build my church; and the gates of hell shall not prevail against it.","And I also say to you that you are Peter, and on this rock I will build My church, and the gates of Hades shall not prevail against it.");
  s(39,28,19,"Go ye therefore, and teach all nations, baptizing them in the name of the Father, and of the Son, and of the Holy Ghost:","Go therefore and make disciples of all the nations, baptizing them in the name of the Father and of the Son and of the Holy Spirit,");
  s(39,28,20,"Teaching them to observe all things whatsoever I have commanded you: and, lo, I am with you alway, even unto the end of the world. Amen.","teaching them to observe all things that I have commanded you; and lo, I am with you always, even to the end of the age. Amen.");
  // MARK
  s(40,10,27,"And Jesus looking upon them saith, With men it is impossible, but not with God: for with God all things are possible.","But Jesus looked at them and said, \"With men it is impossible, but not with God; for with God all things are possible.\"");
  s(40,16,15,"And he said unto them, Go ye into all the world, and preach the gospel to every creature.","And He said to them, \"Go into all the world and preach the gospel to every creature.\"");
  // LUKE
  s(41,1,37,"For with God nothing shall be impossible.","For with God nothing will be impossible.");
  s(41,4,18,"The Spirit of the Lord is upon me, because he hath anointed me to preach the gospel to the poor; he hath sent me to heal the brokenhearted, to preach deliverance to the captives.","The Spirit of the LORD is upon Me, because He has anointed Me to preach the gospel to the poor; He has sent Me to heal the brokenhearted, to proclaim liberty to the captives and recovery of sight to the blind.");
  s(41,6,38,"Give, and it shall be given unto you; good measure, pressed down, and shaken together, and running over.","Give, and it will be given to you: good measure, pressed down, shaken together, and running over will be put into your bosom.");
  // JOHN
  s(42,1,1,"In the beginning was the Word, and the Word was with God, and the Word was God.","In the beginning was the Word, and the Word was with God, and the Word was God.");
  s(42,1,14,"And the Word was made flesh, and dwelt among us, (and we beheld his glory, the glory as of the only begotten of the Father,) full of grace and truth.","And the Word became flesh and dwelt among us, and we beheld His glory, the glory as of the only begotten of the Father, full of grace and truth.");
  s(42,3,16,"For God so loved the world, that he gave his only begotten Son, that whosoever believeth in him should not perish, but have everlasting life.","For God so loved the world that He gave His only begotten Son, that whoever believes in Him should not perish but have everlasting life.");
  s(42,3,17,"For God sent not his Son into the world to condemn the world; but that the world through him might be saved.","For God did not send His Son into the world to condemn the world, but that the world through Him might be saved.");
  s(42,8,32,"And ye shall know the truth, and the truth shall make you free.","And you shall know the truth, and the truth shall make you free.");
  s(42,10,10,"The thief cometh not, but for to steal, and to kill, and to destroy: I am come that they might have life, and that they might have it more abundantly.","The thief does not come except to steal, and to kill, and to destroy. I have come that they may have life, and that they may have it more abundantly.");
  s(42,11,35,"Jesus wept.","Jesus wept.");
  s(42,14,1,"Let not your heart be troubled: ye believe in God, believe also in me.","Let not your heart be troubled; you believe in God, believe also in Me.");
  s(42,14,2,"In my Father's house are many mansions: if it were not so, I would have told you. I go to prepare a place for you.","In My Father's house are many mansions; if it were not so, I would have told you. I go to prepare a place for you.");
  s(42,14,3,"And if I go and prepare a place for you, I will come again, and receive you unto myself; that where I am, there ye may be also.","And if I go and prepare a place for you, I will come again and receive you to Myself; that where I am, there you may be also.");
  s(42,14,6,"Jesus saith unto him, I am the way, the truth, and the life: no man cometh unto the Father, but by me.","Jesus said to him, \"I am the way, the truth, and the life. No one comes to the Father except through Me.\"");
  s(42,14,27,"Peace I leave with you, my peace I give unto you: not as the world giveth, give I unto you. Let not your heart be troubled, neither let it be afraid.","Peace I leave with you, My peace I give to you; not as the world gives do I give to you. Let not your heart be troubled, neither let it be afraid.");
  s(42,15,13,"Greater love hath no man than this, that a man lay down his life for his friends.","Greater love has no one than this, than to lay down one's life for his friends.");
  // ACTS
  s(43,1,8,"But ye shall receive power, after that the Holy Ghost is come upon you: and ye shall be witnesses unto me both in Jerusalem, and in all Judaea, and in Samaria, and unto the uttermost part of the earth.","But you shall receive power when the Holy Spirit has come upon you; and you shall be witnesses to Me in Jerusalem, and in all Judea and Samaria, and to the end of the earth.");
  s(43,2,38,"Then Peter said unto them, Repent, and be baptized every one of you in the name of Jesus Christ for the remission of sins, and ye shall receive the gift of the Holy Ghost.","Then Peter said to them, \"Repent, and let every one of you be baptized in the name of Jesus Christ for the remission of sins; and you shall receive the gift of the Holy Spirit.\"");
  // ROMANS
  s(44,3,23,"For all have sinned, and come short of the glory of God;","for all have sinned and fall short of the glory of God,");
  s(44,5,8,"But God commendeth his love toward us, in that, while we were yet sinners, Christ died for us.","But God demonstrates His own love toward us, in that while we were still sinners, Christ died for us.");
  s(44,6,23,"For the wages of sin is death; but the gift of God is eternal life through Jesus Christ our Lord.","For the wages of sin is death, but the gift of God is eternal life in Christ Jesus our Lord.");
  s(44,8,1,"There is therefore now no condemnation to them which are in Christ Jesus.","There is therefore now no condemnation to those who are in Christ Jesus.");
  s(44,8,28,"And we know that all things work together for good to them that love God, to them who are the called according to his purpose.","And we know that all things work together for good to those who love God, to those who are the called according to His purpose.");
  s(44,8,38,"For I am persuaded, that neither death, nor life, nor angels, nor principalities, nor powers, nor things present, nor things to come,","For I am persuaded that neither death nor life, nor angels nor principalities nor powers, nor things present nor things to come,");
  s(44,8,39,"Nor height, nor depth, nor any other creature, shall be able to separate us from the love of God, which is in Christ Jesus our Lord.","nor height nor depth, nor any other created thing, shall be able to separate us from the love of God which is in Christ Jesus our Lord.");
  s(44,10,9,"That if thou shalt confess with thy mouth the Lord Jesus, and shalt believe in thine heart that God hath raised him from the dead, thou shalt be saved.","that if you confess with your mouth the Lord Jesus and believe in your heart that God has raised Him from the dead, you will be saved.");
  s(44,10,13,"For whosoever shall call upon the name of the Lord shall be saved.","For \"whoever calls on the name of the LORD shall be saved.\"");
  s(44,12,1,"I beseech you therefore, brethren, by the mercies of God, that ye present your bodies a living sacrifice, holy, acceptable unto God, which is your reasonable service.","I beseech you therefore, brethren, by the mercies of God, that you present your bodies a living sacrifice, holy, acceptable to God, which is your reasonable service.");
  s(44,12,2,"And be not conformed to this world: but be ye transformed by the renewing of your mind.","And do not be conformed to this world, but be transformed by the renewing of your mind, that you may prove what is that good and acceptable and perfect will of God.");
  // 1 CORINTHIANS
  s(45,13,4,"Charity suffereth long, and is kind; charity envieth not; charity vaunteth not itself, is not puffed up,","Love suffers long and is kind; love does not envy; love does not parade itself, is not puffed up;");
  s(45,13,13,"And now abideth faith, hope, charity, these three; but the greatest of these is charity.","And now abide faith, hope, love, these three; but the greatest of these is love.");
  s(45,15,55,"O death, where is thy sting? O grave, where is thy victory?","\"O Death, where is your sting? O Hades, where is your victory?\"");
  // 2 CORINTHIANS
  s(46,5,17,"Therefore if any man be in Christ, he is a new creature: old things are passed away; behold, all things are become new.","Therefore, if anyone is in Christ, he is a new creation; old things have passed away; behold, all things have become new.");
  s(46,12,9,"And he said unto me, My grace is sufficient for thee: for my strength is made perfect in weakness.","And He said to me, \"My grace is sufficient for you, for My strength is made perfect in weakness.\"");
  // GALATIANS
  s(47,2,20,"I am crucified with Christ: nevertheless I live; yet not I, but Christ liveth in me.","I have been crucified with Christ; it is no longer I who live, but Christ lives in me.");
  s(47,5,22,"But the fruit of the Spirit is love, joy, peace, longsuffering, gentleness, goodness, faith,","But the fruit of the Spirit is love, joy, peace, longsuffering, kindness, goodness, faithfulness,");
  s(47,5,23,"Meekness, temperance: against such there is no law.","gentleness, self-control. Against such there is no law.");
  s(47,6,9,"And let us not be weary in well doing: for in due season we shall reap, if we faint not.","And let us not grow weary while doing good, for in due season we shall reap if we do not lose heart.");
  // EPHESIANS
  s(48,2,8,"For by grace are ye saved through faith; and that not of yourselves: it is the gift of God:","For by grace you have been saved through faith, and that not of yourselves; it is the gift of God,");
  s(48,2,9,"Not of works, lest any man should boast.","not of works, lest anyone should boast.");
  s(48,3,20,"Now unto him that is able to do exceeding abundantly above all that we ask or think.","Now to Him who is able to do exceedingly abundantly above all that we ask or think, according to the power that works in us,");
  s(48,6,11,"Put on the whole armour of God, that ye may be able to stand against the wiles of the devil.","Put on the whole armor of God, that you may be able to stand against the wiles of the devil.");
  // PHILIPPIANS
  s(49,4,6,"Be careful for nothing; but in every thing by prayer and supplication with thanksgiving let your requests be made known unto God.","Be anxious for nothing, but in everything by prayer and supplication, with thanksgiving, let your requests be made known to God;");
  s(49,4,7,"And the peace of God, which passeth all understanding, shall keep your hearts and minds through Christ Jesus.","and the peace of God, which surpasses all understanding, will guard your hearts and minds through Christ Jesus.");
  s(49,4,13,"I can do all things through Christ which strengtheneth me.","I can do all things through Christ who strengthens me.");
  s(49,4,19,"But my God shall supply all your need according to his riches in glory by Christ Jesus.","And my God shall supply all your need according to His riches in glory by Christ Jesus.");
  // 2 TIMOTHY
  s(54,1,7,"For God hath not given us the spirit of fear; but of power, and of love, and of a sound mind.","For God has not given us a spirit of fear, but of power and of love and of a sound mind.");
  s(54,3,16,"All scripture is given by inspiration of God, and is profitable for doctrine, for reproof, for correction, for instruction in righteousness:","All Scripture is given by inspiration of God, and is profitable for doctrine, for reproof, for correction, for instruction in righteousness,");
  // HEBREWS
  s(57,4,12,"For the word of God is quick, and powerful, and sharper than any twoedged sword.","For the word of God is living and powerful, and sharper than any two-edged sword.");
  s(57,11,1,"Now faith is the substance of things hoped for, the evidence of things not seen.","Now faith is the substance of things hoped for, the evidence of things not seen.");
  s(57,11,6,"But without faith it is impossible to please him.","But without faith it is impossible to please Him, for he who comes to God must believe that He is, and that He is a rewarder of those who diligently seek Him.");
  s(57,12,1,"Wherefore seeing we also are compassed about with so great a cloud of witnesses, let us lay aside every weight, and the sin which doth so easily beset us.","Therefore we also, since we are surrounded by so great a cloud of witnesses, let us lay aside every weight, and the sin which so easily ensnares us.");
  s(57,13,8,"Jesus Christ the same yesterday, and to day, and for ever.","Jesus Christ is the same yesterday, today, and forever.");
  // JAMES
  s(58,1,5,"If any of you lack wisdom, let him ask of God, that giveth to all men liberally, and upbraideth not; and it shall be given him.","If any of you lacks wisdom, let him ask of God, who gives to all liberally and without reproach, and it will be given to him.");
  s(58,4,7,"Submit yourselves therefore to God. Resist the devil, and he will flee from you.","Therefore submit to God. Resist the devil and he will flee from you.");
  // 1 PETER
  s(59,5,7,"Casting all your care upon him; for he careth for you.","casting all your care upon Him, for He cares for you.");
  // 1 JOHN
  s(61,1,9,"If we confess our sins, he is faithful and just to forgive us our sins, and to cleanse us from all unrighteousness.","If we confess our sins, He is faithful and just to forgive us our sins and to cleanse us from all unrighteousness.");
  s(61,4,8,"He that loveth not knoweth not God; for God is love.","He who does not love does not know God, for God is love.");
  s(61,4,19,"We love him, because he first loved us.","We love Him because He first loved us.");
  // REVELATION
  s(65,3,20,"Behold, I stand at the door, and knock: if any man hear my voice, and open the door, I will come in to him, and will sup with him, and he with me.","Behold, I stand at the door and knock. If anyone hears My voice and opens the door, I will come in to him and dine with him, and he with Me.");
  s(65,21,4,"And God shall wipe away all tears from their eyes; and there shall be no more death, neither sorrow, nor crying, neither shall there be any more pain: for the former things are passed away.","And God will wipe away every tear from their eyes; there shall be no more death, nor sorrow, nor crying. There shall be no more pain, for the former things have passed away.");
  s(65,22,13,"I am Alpha and Omega, the beginning and the end, the first and the last.","I am the Alpha and the Omega, the Beginning and the End, the First and the Last.");
}

// ─── JSON LOADER ─────────────────────────────────────────────────────────────
// Called at startup by main process via IPC — loads data/kjv.json into memory.
// Format expected: [{"b":1,"c":1,"v":1,"t":"verse text"}, ...]
// b is 1-indexed (1=Genesis), we store 0-indexed internally.
function loadFromJSON(jsonData, translation) {
  _ensureStore(translation);
  const store = STORES[translation];
  store.clear();
  let count = 0;
  for (const row of jsonData) {
    const key = `${row.b - 1}:${row.c}:${row.v}`;
    store.set(key, row.t);
    count++;
  }
  console.log(`[BibleDB] Loaded ${count} ${translation} verses`);
  _dbStatus = 'full';
  _dbLoaded = true;
  return count;
}

// ─── LOOKUP ENGINE ────────────────────────────────────────────────────────────
function bookNameToIndex(name) {
  if (!name) return null;
  const n = name.trim().toLowerCase().replace(/[^a-z0-9]/g, '');
  if (ABBREV[n] !== undefined) return ABBREV[n];
  for (let i = 0; i < ALL_BOOKS.length; i++) {
    const b = ALL_BOOKS[i].toLowerCase().replace(/[^a-z0-9]/g, '');
    if (b === n || (b.startsWith(n) && n.length >= 3)) return i;
  }
  return null;
}

function indexToBookName(idx) { return ALL_BOOKS[idx] || null; }

function getVerse(book, chapter, verse, translation = 'KJV') {
  const idx = typeof book === 'number' ? book : bookNameToIndex(book);
  if (idx === null) return null;
  const key = `${idx}:${chapter}:${verse}`;
  const store = STORES[translation] || STORES.KJV;
  return store.get(key) || STORES.KJV.get(key) || null;
}

function getChapter(book, chapter, translation = 'KJV') {
  const idx = typeof book === 'number' ? book : bookNameToIndex(book);
  if (idx === null) return [];

  const store = STORES[translation] || STORES.KJV;
  const prefix = `${idx}:${chapter}:`;
  const stored = new Map();

  // Gather stored verses for this chapter
  for (const [key, text] of store.entries()) {
    if (key.startsWith(prefix)) {
      const vn = parseInt(key.split(':')[2]);
      stored.set(vn, text);
    }
  }
  // KJV fallback
  if (store !== STORES.KJV) {
    for (const [key, text] of STORES.KJV.entries()) {
      if (key.startsWith(prefix)) {
        const vn = parseInt(key.split(':')[2]);
        if (!stored.has(vn)) stored.set(vn, text);
      }
    }
  }

  // Get authoritative verse count
  const maxV = (VERSE_COUNTS[idx] && VERSE_COUNTS[idx][chapter - 1]) || 0;
  if (maxV === 0 && stored.size === 0) return [];

  const total = Math.max(maxV, ...stored.keys());
  const results = [];

  for (let vn = 1; vn <= total; vn++) {
    if (stored.has(vn)) {
      results.push({ verse: vn, text: stored.get(vn) });
    } else if (_dbStatus === 'full') {
      // Full DB loaded but verse missing — shouldn't happen
      results.push({ verse: vn, text: `[${indexToBookName(idx)} ${chapter}:${vn}]` });
    } else {
      // Fallback mode: show placeholder so chapter navigation always works
      results.push({ verse: vn,
        text: `[${indexToBookName(idx)} ${chapter}:${vn} — upload kjv.json in Settings › Bible Versions to load full text]`
      });
    }
  }
  return results;
}

function getAvailableChapters(book) {
  const idx = typeof book === 'number' ? book : bookNameToIndex(book);
  if (idx === null || !VERSE_COUNTS[idx]) return [];
  return VERSE_COUNTS[idx].map((_, i) => i + 1);
}

function parseReference(text) {
  if (!text) return null;
  text = text.trim();
  const mRange = text.match(/^((?:\d\s?)?[a-zA-Z]+(?:\s[a-zA-Z]+)?)\s+(\d+)[:\s.]+(\d+)\s*[-–—]\s*(\d+)$/i);
  if (mRange) {
    const idx = bookNameToIndex(mRange[1]);
    if (idx !== null) {
      const startV = +mRange[3], endV = +mRange[4];
      if (endV > startV) return { bookIdx: idx, book: indexToBookName(idx), chapter: +mRange[2], verse: startV, endVerse: endV };
      if (endV === startV) return { bookIdx: idx, book: indexToBookName(idx), chapter: +mRange[2], verse: startV };
    }
  }
  const m = text.match(/^((?:\d\s?)?[a-zA-Z]+(?:\s[a-zA-Z]+)?)\s+(\d+)[:\s.]+(\d+)$/i);
  if (m) {
    const idx = bookNameToIndex(m[1]);
    if (idx !== null) return { bookIdx: idx, book: indexToBookName(idx), chapter: +m[2], verse: +m[3] };
  }
  const spoken = text.match(/^((?:\d\s?)?[a-zA-Z]+(?:\s[a-zA-Z]+)?)\s+chapter\s+(\w+)\s+verse\s+(\w+)$/i);
  if (spoken) {
    const words = { one:1,two:2,three:3,four:4,five:5,six:6,seven:7,eight:8,nine:9,ten:10,
      eleven:11,twelve:12,thirteen:13,fourteen:14,fifteen:15,sixteen:16,seventeen:17,
      eighteen:18,nineteen:19,twenty:20,thirty:30,forty:40,fifty:50 };
    const idx = bookNameToIndex(spoken[1]);
    const ch  = parseInt(spoken[2])  || words[spoken[2].toLowerCase()];
    const vs  = parseInt(spoken[3])  || words[spoken[3].toLowerCase()];
    if (idx !== null && ch && vs) return { bookIdx: idx, book: indexToBookName(idx), chapter: ch, verse: vs };
  }
  return null;
}

function detectReferences(text) {
  const results = [];
  const pattern = /\b((?:\d\s?)?[A-Z][a-zA-Z]+(?:\s[A-Za-z]+)?)\s+(\d+)[:\s.]+(\d+)\b/g;
  let m;
  while ((m = pattern.exec(text)) !== null) {
    const idx = bookNameToIndex(m[1]);
    if (idx !== null) {
      const book = indexToBookName(idx);
      results.push({ bookIdx: idx, book, chapter: +m[2], verse: +m[3],
        ref: `${book} ${m[2]}:${m[3]}`, confidence: 0.95, type: 'direct' });
    }
  }

  const mergedPattern = /\b((?:\d\s?)?[A-Z][a-zA-Z]+(?:\s[A-Za-z]+)?)\s+(\d{2,3})\b/g;
  let mm;
  while ((mm = mergedPattern.exec(text)) !== null) {
    const idx = bookNameToIndex(mm[1]);
    if (idx === null) continue;
    const book = indexToBookName(idx);
    const num = mm[2];
    const ch = +num;
    if (getVerse(book, ch, 1)) continue;
    for (let s = 1; s < num.length; s++) {
      const tryC = parseInt(num.slice(0, s));
      const tryV = parseInt(num.slice(s));
      if (!tryC || !tryV) continue;
      if (getVerse(book, tryC, tryV)) {
        if (!results.some(r => r.book === book && r.chapter === tryC && r.verse === tryV)) {
          results.push({ bookIdx: idx, book, chapter: tryC, verse: tryV,
            ref: `${book} ${tryC}:${tryV}`, confidence: 0.93, type: 'direct' });
        }
        break;
      }
    }
  }

  return results;
}

const _STOP_WORDS = new Set([
  'the','and','of','to','in','a','is','that','for','it','with','was','on','are',
  'be','this','have','from','or','had','by','but','not','you','all','can','her',
  'his','one','our','out','they','we','were','which','will','an','each','she',
  'do','how','if','my','no','he','me','us','so','up','him','has','its','may',
  'who','did','get','has','let','say','too','use','been','come','into','made',
  'than','them','then','what','when','your','also','just','more','some','only',
  'i','shall','unto','thy','thee','thou','thine','ye','upon','hath','doth',
  'art','am','as','at','nor'
]);

function _normalizeForSearch(text) {
  return text.toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function _extractKeywords(text) {
  return _normalizeForSearch(text)
    .split(' ')
    .filter(w => w.length > 2 && !_STOP_WORDS.has(w));
}

function _buildNgrams(words, n) {
  const ngrams = [];
  for (let i = 0; i <= words.length - n; i++) {
    ngrams.push(words.slice(i, i + n).join(' '));
  }
  return ngrams;
}

const _VERSE_INDEX = new Map();
let _indexBuiltFor = null;

function _buildVerseIndex(translation) {
  const store = STORES[translation] || STORES.KJV;
  _VERSE_INDEX.clear();
  for (const [key, text] of store.entries()) {
    if (!text || text.startsWith('\uD83D\uDCE5')) continue;
    const words = _extractKeywords(text);
    if (words.length < 3) continue;
    const wordSet = new Set(words);
    const bigrams = new Set(_buildNgrams(words, 2));
    const trigrams = words.length >= 3 ? new Set(_buildNgrams(words, 3)) : new Set();
    _VERSE_INDEX.set(key, { text, wordSet, bigrams, trigrams, wordCount: wordSet.size });
  }
  _indexBuiltFor = translation + ':' + store.size;
}

function _ensureIndex(translation) {
  const store = STORES[translation] || STORES.KJV;
  const cacheKey = translation + ':' + store.size;
  if (_indexBuiltFor !== cacheKey) _buildVerseIndex(translation);
}

function searchVerses(query, translation = 'KJV', limit = 15) {
  const queryWords = _extractKeywords(query);
  if (queryWords.length < 2) return [];

  _ensureIndex(translation);

  const queryBigrams = _buildNgrams(queryWords, 2);
  const queryTrigrams = _buildNgrams(queryWords, 3);
  const querySet = new Set(queryWords);

  const results = [];

  for (const [key, cached] of _VERSE_INDEX.entries()) {
    let matchCount = 0;
    for (const w of querySet) {
      if (cached.wordSet.has(w)) matchCount++;
    }
    if (matchCount < 2) continue;

    let score = matchCount * 2;

    const overlapRatio = matchCount / Math.min(querySet.size, cached.wordCount);
    if (overlapRatio >= 0.6) score += 5;
    else if (overlapRatio >= 0.4) score += 3;

    let bigramMatch = 0;
    for (const bg of queryBigrams) {
      if (cached.bigrams.has(bg)) bigramMatch++;
    }
    score += bigramMatch * 3;

    if (queryTrigrams.length > 0) {
      for (const tg of queryTrigrams) {
        if (cached.trigrams.has(tg)) score += 5;
      }
    }

    for (const w of queryWords) {
      if (w.length >= 6 && cached.wordSet.has(w)) score += 1;
    }

    if (score >= 4) {
      const [bIdx, ch, vNum] = key.split(':').map(Number);
      results.push({ book: indexToBookName(bIdx), chapter: ch, verse: vNum,
        ref: `${indexToBookName(bIdx)} ${ch}:${vNum}`, text: cached.text, score });
    }
  }
  return results.sort((a, b) => b.score - a.score).slice(0, limit);
}

function getAllRefs() {
  const refs = [];
  for (const key of STORES.KJV.keys()) {
    const [bIdx, ch, vNum] = key.split(':').map(Number);
    refs.push({ book: indexToBookName(bIdx), chapter: ch, verse: vNum,
      ref: `${indexToBookName(bIdx)} ${ch}:${vNum}` });
  }
  return refs;
}

function getDbStatus() {
  return {
    status: _dbStatus,
    verseCount: STORES.KJV.size,
    translations: TRANSLATIONS.filter(t => STORES[t] && STORES[t].size > 0)
  };
}

// ─── INIT ─────────────────────────────────────────────────────────────────────
_seed(); // Always load the 223 built-in key verses first

// Request JSON data from main process (Electron IPC) — loads ALL translations found in data/
if (window.electronAPI && window.electronAPI.loadBibleData) {
  window.electronAPI.loadBibleData().then(data => {
    if (data && typeof data === 'object') {
      let loaded = 0;
      for (const [id, verses] of Object.entries(data)) {
        if (Array.isArray(verses) && verses.length > 10) {
          loadFromJSON(verses, id);
          loaded++;
        }
      }
      if (loaded > 0) {
        console.log(`[BibleDB] Loaded ${loaded} translation(s):`, Object.keys(data).join(', '));
      } else {
        console.log('[BibleDB] Running with built-in verses. Upload a Bible JSON in Settings › Bible Versions.');
      }
    }
  }).catch(e => {
    console.log('[BibleDB] JSON load skipped:', e.message);
  });
}

window.BibleDB = {
  CANON, ALL_BOOKS, TRANSLATIONS,
  getVerse, getChapter, getAvailableChapters,
  parseReference, detectReferences, searchVerses,
  getAllRefs, bookNameToIndex, indexToBookName, loadFromJSON, getDbStatus,
  getChapters: getAvailableChapters,
  normalizeBook: (n) => indexToBookName(bookNameToIndex(n)),
};
