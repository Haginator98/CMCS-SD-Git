class ParticleSystem {
  constructor() {
    this.particles = [];
  }

  create(type, x, y, color, count = 10) {
    for (let i = 0; i < count; i++) {
      this.particles.push({
        type,
        x,
        y,
        vx: Math.random() * 4 - 2,
        vy: Math.random() * 4 - 2,
        size: Math.random() * 5 + 2,
        color,
        life: 100,
        decay: Math.random() * 0.5 + 0.5
      });
    }
  }

  update() {
    for (let i = this.particles.length - 1; i >= 0; i--) {
      const p = this.particles[i];
      p.x += p.vx;
      p.y += p.vy;
      p.life -= p.decay;
      
      if (p.life <= 0) {
        this.particles.splice(i, 1);
      }
    }
  }

  draw(ctx) {
    ctx.save();
    for (const p of this.particles) {
      ctx.globalAlpha = p.life / 100;
      ctx.fillStyle = p.color;
      ctx.beginPath();
      ctx.arc(p.x, p.y, p.size, 0, Math.PI * 2);
      ctx.fill();
    }
    ctx.restore();
  }
}

const particles = new ParticleSystem();