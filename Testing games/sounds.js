class SoundManager {
  constructor() {
    this.sounds = {};
    this.muted = false;
  }

  load(name, src) {
    this.sounds[name] = new Audio(src);
  }

  play(name) {
    if (!this.muted && this.sounds[name]) {
      this.sounds[name].currentTime = 0;
      this.sounds[name].play();
    }
  }

  toggleMute() {
    this.muted = !this.muted;
  }
}

const soundManager = new SoundManager();

// Load default sounds
soundManager.load('tower_attack', 'https://assets.mixkit.co/sfx/preview/mixkit-short-laser-gun-shot-1670.mp3');
soundManager.load('enemy_death', 'https://assets.mixkit.co/sfx/preview/mixkit-arcade-game-explosion-2759.mp3');
soundManager.load('game_over', 'https://assets.mixkit.co/sfx/preview/mixkit-retro-arcade-lose-2027.mp3');
soundManager.load('round_start', 'https://assets.mixkit.co/sfx/preview/mixkit-unlock-game-notification-253.mp3');

export { soundManager };