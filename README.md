
# Nanop's Chill Music Player

Super simple music player for windows that barely takes up any ram with built in ambience. This one is special, because i made this in just under 6 hours with VB6.

If you think having ambience was a random feature, because it is. I initially wasn't planning on releasing this, so i added whatever i wanted lol



## Installation And Usage
All you have to do, is install the 400kb exe file, and put it inside a separate folder of it's own. To actually use it, all you gotta do is just put a music folder next to it, and set it up like this:
```
crappymusicplayeroffofgithub
├── NanopsChillMusicPlayer.exe
└── music
    ├── yoursong.mp3
    ├── ambience
    │   └── rain.mp3
    │   └── mommyasmr.mp3
    ├── yourplaylist
    │   └── yourothersong.mp3
    └── yoursecondplaylist
        └── namesongsproperly.mp3
```

If you're too lazy to set this up, don't worry, theres also a zip version download where the folders are setup for you, and it actually also includes rain.mp3.

ˡᵃᶻʸ ᵃʰʰ
## Inspiration

I created this project because of two concepts my last remaining braincells encountered

- "Hey. It would be nice if my music had ambience, like rain"
- "Why the freak is my music player taking 500mb of ram"


## Technologies
Literally 90's tech. This solves problem #2. Im using it right now, it's maxxing out on 16mb of ram on my system. This is because it was built with Visual Basic 6.0, which was built for older machines which didn't have much ram to work with. Initially i wanted to use VB6 as like a prototyping tool before i switched to something more modern, like C# Forms, but i got too carried away.

I had to learn Visual Basic in that same 6 hours while making it. Side effect of using 90's tech, is that this probably runs on windows xp on a good day lol

---
ps. It's technically just windows media player with a bunch of fancy clothes on but windows media player could never look as dapper and have cool features as this


## Building locally

First off, you need a copy of VB6. I won't link any here because i probably got a virus while installing it. Getting VB6 to actually run in the first place as well is also a challenge, so good luck with that.

If you actually managed to get VB6 running, you can go ahead and extract the files into a folder, and open the one called "lofi.vbp". Once you're there, go ahead and start to cry while you read the code line by line and realize that it's not actually maintainable. Remember, it was initially a prototype while at the same time i was learning the language. In addition to that, the VB6 IDE sucks ass. Good luck.


## Known bugs

- Sometimes when changing songs for the first time, the player goes into the foreground and focuses itself. No clue why it does that, but it only does that once. So do i even fix it?????
- Audio can get delayed when you scrub too fast on lengthy songs. This sucks on songs about 4mins in length, because you have to sit through atleast 5 seconds of audio scrubbing.

