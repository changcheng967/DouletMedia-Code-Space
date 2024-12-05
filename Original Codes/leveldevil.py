import pygame
import random
import sys
import requests  # Import the requests library for HTTP requests

# Initialize pygame
pygame.init()

# Screen dimensions
WIDTH, HEIGHT = 600, 600
PLAYER_SIZE = 20
OBSTACLE_SIZE = 30
CELL_SIZE = 20

# Colors
BLACK = (0, 0, 0)
WHITE = (255, 255, 255)
GREEN = (0, 255, 0)
RED = (255, 0, 0)

# Game settings
NUM_OBSTACLES = 5  # Starting number of obstacles
LEVEL_TIME_LIMIT = 3000  # Level duration in milliseconds (3 seconds)
obstacle_speed = 15  # Speed at which obstacles fall
player_speed = 10

# Webhook URL (your Discord webhook URL)
WEBHOOK_URL = "https://discord.com/api/webhooks/1309688711063011398/sgbmD1fGWsVjFDBdGyp_gMw1ZxEWvbv-NoYbdYKC4_IZOtwGIH62TQQlwWrhh0PHFudZ"

# Initialize screen
screen = pygame.display.set_mode((WIDTH, HEIGHT))
pygame.display.set_caption("Level Devil")

# Clock for controlling game speed
clock = pygame.time.Clock()

# Function to send score and level to Discord
def send_score_to_discord(score, level, username="Player"):
    message = {
        "content": f"**Game Over!**\n**Score**: {score}\n**Level**: {level}\n**Player**: {username}",
        "embeds": [{
            "title": "Game Over",
            "description": f"{username} scored {score} points at level {level}!",
            "color": 16711680  # Red color
        }]
    }
    try:
        response = requests.post(WEBHOOK_URL, json=message)
        if response.status_code == 204:
            print("Score and level sent to Discord successfully!")
        else:
            print(f"Failed to send score and level to Discord: {response.status_code}")
    except Exception as e:
        print(f"Error sending score and level to Discord: {e}")

# Function to get the username
def get_username():
    font = pygame.font.Font(None, 36)
    input_box = pygame.Rect(WIDTH // 2 - 100, HEIGHT // 2, 200, 50)
    color_inactive = pygame.Color('lightskyblue3')
    color_active = pygame.Color('dodgerblue2')
    color = color_inactive
    active = False
    text = ''
    clock = pygame.time.Clock()

    while True:
        screen.fill(BLACK)
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                pygame.quit()
                sys.exit()

            if event.type == pygame.MOUSEBUTTONDOWN:
                if input_box.collidepoint(event.pos):
                    active = not active
                else:
                    active = False

            if event.type == pygame.KEYDOWN:
                if active:
                    if event.key == pygame.K_RETURN:
                        return text if text else "Player"
                    elif event.key == pygame.K_BACKSPACE:
                        text = text[:-1]
                    else:
                        text += event.unicode

        txt_surface = font.render(text, True, color)
        width = max(200, txt_surface.get_width() + 10)
        input_box.w = width
        screen.blit(txt_surface, (input_box.x + 5, input_box.y + 5))
        pygame.draw.rect(screen, color, input_box, 2)

        prompt = font.render("Enter your username (press Enter to start):", True, WHITE)
        screen.blit(prompt, (WIDTH // 2 - prompt.get_width() // 2, HEIGHT // 4))

        pygame.display.flip()
        clock.tick(30)

# Function to display text on screen
def display_text(text, font, color, position):
    text_surface = font.render(text, True, color)
    screen.blit(text_surface, position)

# Function to create obstacles
def create_obstacles(num_obstacles):
    obstacles = []
    for _ in range(num_obstacles):
        x = random.randint(0, WIDTH - OBSTACLE_SIZE)
        y = random.randint(-100, -OBSTACLE_SIZE)
        obstacles.append([x, y])
    return obstacles

# Function to update the obstacles falling down
def update_obstacles(obstacles):
    for obstacle in obstacles:
        obstacle[1] += obstacle_speed  # Make obstacles fall faster
        if obstacle[1] > HEIGHT:  # Reset the obstacle when it moves out of the screen
            obstacle[1] = random.randint(-100, -OBSTACLE_SIZE)
            obstacle[0] = random.randint(0, WIDTH - OBSTACLE_SIZE)
    return obstacles

# Function to check if the player collides with any obstacle
def check_collision(player_rect, obstacles):
    for x, y in obstacles:
        obstacle_rect = pygame.Rect(x, y, OBSTACLE_SIZE, OBSTACLE_SIZE)
        if player_rect.colliderect(obstacle_rect):
            return True
    return False

# Main game function
def game():
    global NUM_OBSTACLES, player_speed, obstacle_speed  # Ensure NUM_OBSTACLES is used as a global variable
    username = get_username()  # Get the username before starting the game

    while True:
        # Initialize game variables
        player_x, player_y = WIDTH // 2, HEIGHT - 2 * PLAYER_SIZE
        score = 0
        level = 1
        obstacles = create_obstacles(NUM_OBSTACLES)
        level_start_time = pygame.time.get_ticks()  # Track time for level duration
        font = pygame.font.Font(None, 36)
        last_score_update_time = pygame.time.get_ticks()  # Track the last time the score was updated

        # Main game loop
        while True:
            screen.fill(BLACK)

            # Event handling
            for event in pygame.event.get():
                if event.type == pygame.QUIT:
                    pygame.quit()
                    sys.exit()

            # Player movement (with boundary checking)
            keys = pygame.key.get_pressed()
            if keys[pygame.K_LEFT] and player_x > 0:
                player_x -= player_speed
            if keys[pygame.K_RIGHT] and player_x < WIDTH - PLAYER_SIZE:
                player_x += player_speed

            # Update obstacles falling down
            obstacles = update_obstacles(obstacles)

            # Check for collisions
            player_rect = pygame.Rect(player_x, player_y, PLAYER_SIZE, PLAYER_SIZE)
            if check_collision(player_rect, obstacles):
                # Display Game Over and score
                screen.fill(BLACK)
                display_text(f"Game Over! Score: {score}", font, RED, (WIDTH // 3, HEIGHT // 2))
                display_text(f"Level: {level}", font, WHITE, (WIDTH // 3, HEIGHT // 2 + 50))
                display_text("Press 'R' to Retry or 'Q' to Quit", font, WHITE, (WIDTH // 3, HEIGHT // 2 + 100))
                pygame.display.flip()

                # Send the score and level to Discord when the game is over
                send_score_to_discord(score, level, username)

                # Wait for the player's response
                waiting_for_retry = True
                while waiting_for_retry:
                    for event in pygame.event.get():
                        if event.type == pygame.QUIT:
                            pygame.quit()
                            sys.exit()
                        if event.type == pygame.KEYDOWN:
                            if event.key == pygame.K_r:
                                waiting_for_retry = False  # Retry the game
                            if event.key == pygame.K_q:
                                pygame.quit()
                                sys.exit()  # Quit the game
                break  # Exit the game loop to restart

            # Update the score every second
            current_time = pygame.time.get_ticks()
            if current_time - last_score_update_time >= 1000:  # 1000ms = 1 second
                score += 1  # Add 1 point for every second survived
                last_score_update_time = current_time  # Update the last score update time

            # Draw the player
            pygame.draw.rect(screen, GREEN, player_rect)

            # Draw obstacles
            for x, y in obstacles:
                pygame.draw.rect(screen, RED, pygame.Rect(x, y, OBSTACLE_SIZE, OBSTACLE_SIZE))

            # Display the score and level
            display_text(f"Score: {score}", font, WHITE, (10, 10))
            display_text(f"Level: {level}", font, WHITE, (WIDTH - 120, 10))

            # Check if the player has survived the level (every 3 seconds)
            if pygame.time.get_ticks() - level_start_time > LEVEL_TIME_LIMIT:  # 3 seconds per level
                level += 1  # Increase the level
                player_speed += 0.5  # Increase player speed
                obstacle_speed += 2  # Increase obstacle fall speed even more
                NUM_OBSTACLES += 1  # Add more obstacles each level
                obstacles = create_obstacles(NUM_OBSTACLES)  # Recreate obstacles with the new count

                level_start_time = pygame.time.get_ticks()  # Reset the level time

            pygame.display.flip()
            clock.tick(60)  # Control the game loop speed

if __name__ == "__main__":
    game()
