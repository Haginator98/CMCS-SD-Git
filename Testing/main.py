# main.py

import pygame
import sys
import math
import random # Needed for path generation

# --- Constants ---
SCREEN_WIDTH = 800
SCREEN_HEIGHT = 600
BACKGROUND_IMAGE_FILE = "background.png"
ENEMY_IMAGE_FILE = "enemy.png"
TOWER_IMAGE_FILE = "tower.png"

# Enemy Stats
ENEMY_SPEED = 2
ENEMY_WIDTH = 40
ENEMY_HEIGHT = 40
ENEMY_HEALTH = 50
ENEMY_GOLD_REWARD = 15      # Gold awarded for defeating one enemy

# Tower Stats & Placement
TOWER_WIDTH = 50
TOWER_HEIGHT = 50
TOWER_RANGE = 150
TOWER_COOLDOWN = 800 # Milliseconds
TOWER_DAMAGE = 12
TOWER_COST = 75             # Cost to place a tower

# Player & Game
PLAYER_STARTING_GOLD = 400

# Colors
PATH_COLOR = (255, 0, 0)          # Red
WAYPOINT_COLOR = (255, 255, 0)    # Yellow
SHOT_LINE_COLOR = (255, 255, 255) # White
UI_TEXT_COLOR = (255, 255, 0)     # Yellow for UI text
UI_INSTRUCT_COLOR = (255, 255, 255) # White for instructions
MENU_BG_COLOR = (30, 30, 30)      # Dark Gray for menu background
MENU_TITLE_COLOR = (200, 200, 255) # Light Blue/Purple for title
MENU_BUTTON_COLOR = (200, 200, 200) # Light Gray for buttons
PLACEMENT_INVALID_COLOR = (255, 0, 0, 150) # Semi-transparent red
PLACEMENT_VALID_COLOR = (0, 255, 0, 150)   # Semi-transparent green
RANGE_PREVIEW_ALPHA = 100 # Alpha for range preview circle color

# Timings & Waves
TIME_BETWEEN_SPAWNS = 500   # Milliseconds between enemy spawns within a wave
TIME_BETWEEN_WAVES = 5000  # Milliseconds delay before starting next wave/round
SHOT_LINE_DURATION = 75   # How long the shot line stays visible (milliseconds)

# Wave Data Structure
ROUNDS_DATA = [
    # Round 1
    [ {'type': 'basic', 'count': 5} ],
    # Round 2
    [ {'type': 'basic', 'count': 8}, {'type': 'basic', 'count': 10} ],
    # Round 3
    [ {'type': 'basic', 'count': 15}, {'type': 'basic', 'count': 5} ],
    # Add more rounds here...
]


# --- Global Variables ---
player_gold = PLAYER_STARTING_GOLD
# app_state controls the overall application: "menu", "game", "game_over"
app_state = "menu"
# game_state is for IN-GAME states: "between_rounds", "wave_active", "placing_tower"
game_state = "between_rounds" # Default state *when game starts*

# Wave/Round Tracking (initialized/reset when game starts)
current_round = 1
current_wave_index = 0
enemies_left_to_spawn = 0
last_spawn_time = 0
wave_complete_time = 0

# Game Objects (initialized/reset when game starts)
all_enemies = []
all_towers = []
enemy_path = None # Will be generated when game starts

# Assets (loaded once)
tower_preview_surface = None
background_image = None
ui_font = None
ui_title_font = None
screen = None
clock = None

# Menu button rects (calculated in menu loop)
start_button_rect = None
quit_button_rect = None


# --- Enemy Class ---
class Enemy:
    def __init__(self, path, image_file):
        global enemy_path # Access the global path variable
        self.path = enemy_path # Use the currently generated path
        if not self.path or len(self.path) == 0:
             print("ERROR: Enemy created with no valid path!")
             # Handle error appropriately - maybe set enemy to inactive?
             self.is_alive = False
             return # Stop initialization if path is bad

        self.max_health = ENEMY_HEALTH
        self.health = ENEMY_HEALTH
        self.is_alive = True
        try:
            original_image_loaded = pygame.image.load(image_file).convert_alpha()
            self.original_image = pygame.transform.scale(original_image_loaded, (ENEMY_WIDTH, ENEMY_HEIGHT))
        except pygame.error as e:
            print(f"Error loading or scaling enemy image: {image_file} - {e}")
            self.original_image = pygame.Surface((ENEMY_WIDTH, ENEMY_HEIGHT))
            self.original_image.fill((128, 0, 128)) # Purple placeholder

        self.image = self.original_image
        self.rect = self.image.get_rect()
        self.position = pygame.Vector2(self.path[0]) # Start at path start
        self.rect.center = self.position
        self.path_index = 0
        self.speed = ENEMY_SPEED

    def move(self):
        if not self.is_alive or not self.path:
            return
        # Check if there's a next waypoint in the path
        if self.path_index < len(self.path) - 1:
            target_waypoint = self.path[self.path_index + 1]
            target_pos = pygame.Vector2(target_waypoint)
            direction = target_pos - self.position
            distance = direction.length()
            # If close enough (or passed), snap to target and advance path index
            if distance <= self.speed:
                self.position = target_pos
                self.path_index += 1
            else:
                # Calculate movement vector (normalized direction * speed)
                movement = direction.normalize() * self.speed
                self.position += movement
            self.rect.center = self.position # Keep rect updated
        else:
            # Reached end - Mark as not alive (or handle player damage)
            # print("Enemy reached end - Despawning.") # Reduce spam
            self.is_alive = False # Treat reaching the end same as dying for now

    def take_damage(self, amount):
        global player_gold # Declare intent to modify the global variable
        if not self.is_alive:
            return
        self.health -= amount
        # print(f"Enemy took {amount} damage, health: {self.health}") # Reduce console spam
        if self.health <= 0:
            self.is_alive = False
            print("Enemy died!")
            player_gold += ENEMY_GOLD_REWARD
            print(f"Player earned {ENEMY_GOLD_REWARD} gold. Total: {player_gold}")

    def draw(self, surface):
        if self.is_alive:
            surface.blit(self.image, self.rect)

# --- Tower Class ---
class Tower:
    def __init__(self, position, image_file):
        self.position = pygame.Vector2(position)
        self.range = TOWER_RANGE
        self.cooldown = TOWER_COOLDOWN
        self.damage = TOWER_DAMAGE
        self.last_shot_time = 0
        self.shot_target_pos = None
        self.shot_visualization_end_time = 0

        try:
            original_image_loaded = pygame.image.load(image_file).convert_alpha()
            self.original_image = pygame.transform.scale(original_image_loaded, (TOWER_WIDTH, TOWER_HEIGHT))
        except pygame.error as e:
            print(f"Error loading or scaling tower image for instance: {image_file} - {e}")
            self.original_image = pygame.Surface((TOWER_WIDTH, TOWER_HEIGHT))
            self.original_image.fill((100, 100, 100)) # Gray placeholder

        self.image = self.original_image
        self.rect = self.image.get_rect(center=self.position)

    def update(self, current_time, enemies):
        target_enemy = None
        min_distance = self.range + 1 # Start with distance > range
        # Find the closest living enemy within range
        for enemy in enemies:
            if enemy.is_alive:
                distance = self.position.distance_to(enemy.position)
                if distance <= self.range:
                    if distance < min_distance:
                        min_distance = distance
                        target_enemy = enemy

        # Check if we have a target and if the cooldown has passed
        if target_enemy is not None and (current_time - self.last_shot_time >= self.cooldown):
            target_enemy.take_damage(self.damage)
            self.last_shot_time = current_time
            # print(f"Tower at {self.position} shooting!") # Reduce spam
            self.shot_target_pos = target_enemy.position # Record target position for drawing
            self.shot_visualization_end_time = current_time + SHOT_LINE_DURATION

    def draw(self, surface, current_time):
        surface.blit(self.image, self.rect)
        # Draw the shot line if it's currently active
        if self.shot_target_pos is not None and current_time < self.shot_visualization_end_time:
             pygame.draw.line(surface, SHOT_LINE_COLOR, self.position, self.shot_target_pos, 2)
        # Clear the target pos if the visualization time has passed to stop drawing
        elif current_time >= self.shot_visualization_end_time:
             self.shot_target_pos = None

# --- Helper Functions ---
def is_placement_valid(position, tower_width, tower_height, existing_towers):
    """Checks if placing a tower at the given position is valid."""
    global screen # Access the screen surface defined globally
    new_tower_rect = pygame.Rect(0, 0, tower_width, tower_height)
    new_tower_rect.center = position

    if not screen or not screen.get_rect().contains(new_tower_rect):
        return False # Check screen bounds

    for tower in existing_towers:
        if new_tower_rect.colliderect(tower.rect):
            return False # Check collision with existing towers

    # TODO: Add path collision check if desired
    return True

def generate_random_path(num_intermediate_waypoints, screen_width, screen_height):
    """Generates a somewhat random path across the screen."""
    print("Generating new random path...")
    path = []
    min_y = 50
    max_y = screen_height - 50
    start_x = 0
    end_x = screen_width

    start_y = random.randint(min_y, max_y)
    path.append((start_x, start_y))

    num_bands = num_intermediate_waypoints + 1
    band_width = screen_width / num_bands
    last_x = start_x
    last_y = start_y

    for i in range(num_intermediate_waypoints):
        band_center_x = (i + 1) * band_width
        waypoint_x = random.randint(int(band_center_x - band_width * 0.4), int(band_center_x + band_width * 0.4))
        waypoint_x = max(last_x + 30, waypoint_x)
        waypoint_x = min(waypoint_x, screen_width - 30)

        max_y_change = 150
        min_waypoint_y = max(min_y, last_y - max_y_change)
        max_waypoint_y = min(max_y, last_y + max_y_change)
        # Ensure min/max are valid before randint
        if min_waypoint_y > max_waypoint_y: min_waypoint_y = max_waypoint_y
        waypoint_y = random.randint(min_waypoint_y, max_waypoint_y)

        path.append((waypoint_x, waypoint_y))
        last_x = waypoint_x
        last_y = waypoint_y

    end_y = random.randint(min_y, max_y)
    path.append((end_x, end_y))

    print(f"Generated path: {path}")
    return path


# --- Initialization --- (Runs Once)
pygame.init()
pygame.font.init()
try:
    ui_font = pygame.font.SysFont(None, 30)
    ui_title_font = pygame.font.SysFont(None, 72)
except Exception as e:
    print(f"Error initializing font: {e}")
    pygame.quit()
    sys.exit()

screen = pygame.display.set_mode((SCREEN_WIDTH, SCREEN_HEIGHT))
pygame.display.set_caption("Haginator TD V0.2")
clock = pygame.time.Clock()

# --- Load Assets --- (Runs Once)
try:
    background_image = pygame.image.load(BACKGROUND_IMAGE_FILE).convert()
except pygame.error as e:
    print(f"FATAL: Error loading background image: {BACKGROUND_IMAGE_FILE} - {e}")
    pygame.quit()
    sys.exit()

try:
    base_tower_img = pygame.image.load(TOWER_IMAGE_FILE).convert_alpha()
    scaled_base_tower_img = pygame.transform.scale(base_tower_img, (TOWER_WIDTH, TOWER_HEIGHT))
    tower_preview_surface = scaled_base_tower_img.copy()
    tower_preview_surface.set_alpha(150)
except pygame.error as e:
     print(f"ERROR: Could not load tower image for preview: {e}. Placement disabled.")
     tower_preview_surface = None


# --- Main Application Loop ---
running = True
while running:

    # --- State: MENU ---
    if app_state == "menu":
        screen.fill(MENU_BG_COLOR)
        mouse_pos = pygame.mouse.get_pos()

        # Title
        title_surface = ui_title_font.render("Haginator TD!", True, MENU_TITLE_COLOR)
        title_rect = title_surface.get_rect(center=(SCREEN_WIDTH // 2, SCREEN_HEIGHT // 3))
        screen.blit(title_surface, title_rect)

        # Buttons
        start_surface = ui_font.render("Start Game", True, MENU_BUTTON_COLOR)
        start_button_rect = start_surface.get_rect(center=(SCREEN_WIDTH // 2, SCREEN_HEIGHT // 2))
        screen.blit(start_surface, start_button_rect)

        quit_surface = ui_font.render("Quit Game", True, MENU_BUTTON_COLOR)
        quit_button_rect = quit_surface.get_rect(center=(SCREEN_WIDTH // 2, SCREEN_HEIGHT // 2 + 50))
        screen.blit(quit_surface, quit_button_rect)

        # Menu Event Handling
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                running = False
            if event.type == pygame.MOUSEBUTTONDOWN:
                if event.button == 1:
                    if start_button_rect and start_button_rect.collidepoint(mouse_pos):
                        print("Starting game...")
                        app_state = "game"
                        # Reset Game Variables
                        player_gold = PLAYER_STARTING_GOLD
                        game_state = "between_rounds"
                        current_round = 1
                        current_wave_index = 0
                        enemies_left_to_spawn = 0
                        last_spawn_time = 0
                        wave_complete_time = pygame.time.get_ticks() # Allow immediate start trigger
                        all_enemies = []
                        all_towers = []
                        # Generate Random Path for the new game
                        enemy_path = generate_random_path(num_intermediate_waypoints=3, screen_width=SCREEN_WIDTH, screen_height=SCREEN_HEIGHT)
                        pygame.display.set_caption("My Tower Defense - Round 1")

                    elif quit_button_rect and quit_button_rect.collidepoint(mouse_pos):
                        print("Quitting game...")
                        running = False

    # --- State: GAME ---
    elif app_state == "game":
        current_time = pygame.time.get_ticks()
        mouse_pos = pygame.mouse.get_pos()

        # Event Handling (In-Game)
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                running = False
            if event.type == pygame.KEYDOWN:
                if event.key == pygame.K_t and game_state != "wave_active":
                     if player_gold >= TOWER_COST:
                         if tower_preview_surface:
                             game_state = "placing_tower"
                         else: print("Preview error.")
                     else: print("Not enough gold.")
                elif event.key == pygame.K_ESCAPE and game_state == "placing_tower":
                     game_state = "between_rounds"
                elif event.key == pygame.K_SPACE and game_state == "between_rounds":
                     # Start Wave Logic
                     time_since_wave_end = current_time - wave_complete_time
                     if time_since_wave_end >= TIME_BETWEEN_WAVES:
                          if current_round - 1 < len(ROUNDS_DATA):
                              current_round_data = ROUNDS_DATA[current_round - 1]
                              if current_wave_index < len(current_round_data):
                                  current_wave_data = current_round_data[current_wave_index]
                                  enemies_left_to_spawn = current_wave_data.get('count', 0)
                                  last_spawn_time = current_time # Ready to spawn
                                  game_state = "wave_active"
                                  print(f"Starting Round {current_round}, Wave {current_wave_index + 1}")
                                  pygame.display.set_caption(f"My Tower Defense - R: {current_round} W: {current_wave_index + 1}")
                              else: print("Error: Invalid wave index.")
                          else: print("Game Complete!") # Already checked in wave completion
                     else:
                          print("Wait for timer or press SPACE later.")

            if event.type == pygame.MOUSEBUTTONDOWN:
                if game_state == "placing_tower":
                    if event.button == 1: # Left
                        if is_placement_valid(mouse_pos, TOWER_WIDTH, TOWER_HEIGHT, all_towers):
                            if player_gold >= TOWER_COST:
                                all_towers.append(Tower(mouse_pos, TOWER_IMAGE_FILE))
                                player_gold -= TOWER_COST
                                game_state = "between_rounds"
                            else: print("Not enough gold.")
                        else: print("Invalid location.")
                    elif event.button == 3: # Right
                        game_state = "between_rounds"

        # Game Logic Updates (Only during active wave)
        if game_state == "wave_active":
            for enemy in all_enemies: enemy.move()
            for tower in all_towers: tower.update(current_time, all_enemies)
            # Spawning
            if enemies_left_to_spawn > 0 and (current_time - last_spawn_time >= TIME_BETWEEN_SPAWNS):
                all_enemies.append(Enemy(enemy_path, ENEMY_IMAGE_FILE))
                enemies_left_to_spawn -= 1
                last_spawn_time = current_time

        # Cleanup Dead Enemies (Always run)
        living_enemies = []
        for enemy in all_enemies:
            if enemy.is_alive: living_enemies.append(enemy)
        all_enemies = living_enemies

        # Wave Completion Check
        if game_state == "wave_active" and enemies_left_to_spawn == 0 and len(all_enemies) == 0:
            print(f"Wave {current_wave_index + 1} Complete!")
            game_state = "between_rounds"
            wave_complete_time = current_time
            current_wave_index += 1
            if current_round - 1 < len(ROUNDS_DATA) and current_wave_index >= len(ROUNDS_DATA[current_round - 1]):
                print(f"Round {current_round} Complete!")
                current_round += 1
                current_wave_index = 0
                if current_round - 1 >= len(ROUNDS_DATA):
                    print("Congratulations! You beat all defined rounds!")
                    app_state = "menu" # Go back to menu after winning

        # --- Drawing (In-Game) ---
        screen.blit(background_image, (0, 0))
        # Draw Path
        if enemy_path and len(enemy_path) >= 2:
             pygame.draw.lines(screen, PATH_COLOR, False, enemy_path, 5)
             # Draw waypoints if desired for debugging
             # for point in enemy_path: pygame.draw.circle(screen, WAYPOINT_COLOR, point, 10)

        # Draw Towers & Enemies
        for tower in all_towers: tower.draw(screen, current_time)
        for enemy in all_enemies: enemy.draw(screen)

        # Draw Placement Preview
        if game_state == "placing_tower" and tower_preview_surface:
            mouse_pos_vec = pygame.Vector2(mouse_pos)
            valid = is_placement_valid(mouse_pos, TOWER_WIDTH, TOWER_HEIGHT, all_towers)
            preview_range_color = PLACEMENT_VALID_COLOR if valid else PLACEMENT_INVALID_COLOR
            if len(preview_range_color) == 3: preview_range_color = (*preview_range_color, RANGE_PREVIEW_ALPHA)
            preview_rect = tower_preview_surface.get_rect(center=mouse_pos)
            screen.blit(tower_preview_surface, preview_rect)
            range_preview_surface = pygame.Surface((TOWER_RANGE * 2, TOWER_RANGE * 2), pygame.SRCALPHA)
            pygame.draw.circle(range_preview_surface, preview_range_color, (TOWER_RANGE, TOWER_RANGE), TOWER_RANGE)
            screen.blit(range_preview_surface, (mouse_pos_vec.x - TOWER_RANGE, mouse_pos_vec.y - TOWER_RANGE))

        # Draw UI
        gold_text_surface = ui_font.render(f"Gold: {player_gold}", True, UI_TEXT_COLOR)
        screen.blit(gold_text_surface, (10, 10))
        # Determine round/wave display text safely
        round_display = current_round if current_round - 1 < len(ROUNDS_DATA) else 'END'
        wave_display = '-'
        if current_round - 1 < len(ROUNDS_DATA):
            if current_wave_index < len(ROUNDS_DATA[current_round - 1]):
                wave_display = current_wave_index + 1
            else:
                 wave_display = 'END' # Round finished

        round_wave_text = f"Round: {round_display} | Wave: {wave_display}"
        round_wave_surface = ui_font.render(round_wave_text, True, UI_INSTRUCT_COLOR)
        screen.blit(round_wave_surface, (SCREEN_WIDTH - round_wave_surface.get_width() - 10, 10))

        if game_state == "wave_active":
             enemies_text = f"Enemies: {len(all_enemies) + enemies_left_to_spawn}"
             enemies_surface = ui_font.render(enemies_text, True, UI_INSTRUCT_COLOR)
             screen.blit(enemies_surface, (SCREEN_WIDTH - enemies_surface.get_width() - 10, 40))

        y_offset = 40
        if game_state == "placing_tower":
             cost_text = f"Cost: {TOWER_COST} | Left-Click: Place | Right-Click/ESC: Cancel"
             cost_surf = ui_font.render(cost_text, True, UI_INSTRUCT_COLOR)
             screen.blit(cost_surf, (10, y_offset))
        elif game_state == "between_rounds":
             time_since_wave_end = current_time - wave_complete_time
             time_until_next = max(0, TIME_BETWEEN_WAVES - time_since_wave_end)
             start_prompt_text = ""
             if current_round - 1 >= len(ROUNDS_DATA):
                 start_prompt_text = "All rounds complete! (Return to Menu soon)" # Should already transition state
             elif time_until_next > 0:
                  start_prompt_text = f"Next wave in: {time_until_next / 1000:.1f}s (SPACE to skip)"
             else:
                  start_prompt_text = "Press SPACE to start next wave!"
             start_wave_surf = ui_font.render(start_prompt_text, True, UI_INSTRUCT_COLOR)
             screen.blit(start_wave_surf, (10, y_offset))
             y_offset += 30
             place_instruct_surf = ui_font.render("Press 'T' to place tower", True, UI_INSTRUCT_COLOR)
             screen.blit(place_instruct_surf, (10, y_offset))
        elif game_state == "wave_active":
             place_instruct_surf = ui_font.render("Cannot place towers during active wave", True, (200, 200, 200))
             screen.blit(place_instruct_surf, (10, y_offset))

    # --- Update Display --- (Common)
    pygame.display.flip()

    # --- Control Framerate --- (Common)
    clock.tick(60)


# --- Cleanup ---
print("Exiting game.")
pygame.font.quit()
pygame.quit()
sys.exit()