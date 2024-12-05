import pygame
import random
import sys
import requests

# Initialize pygame
pygame.init()

# Screen dimensions
WIDTH, HEIGHT = 600, 600
CELL_SIZE = 20

# Colors
BLACK = (0, 0, 0)
WHITE = (255, 255, 255)
GREEN = (0, 255, 0)
RED = (255, 0, 0)
BLUE = (0, 0, 255)
YELLOW = (255, 255, 0)

# Directions
UP = (0, -1)
DOWN = (0, 1)
LEFT = (-1, 0)
RIGHT = (1, 0)

# Initialize screen
screen = pygame.display.set_mode((WIDTH, HEIGHT))
pygame.display.set_caption("贪吃蛇 - Extreme Mode")

# Clock for controlling game speed
clock = pygame.time.Clock()

# Webhook URL (your Discord webhook URL)
WEBHOOK_URL = "https://discord.com/api/webhooks/1309688711063011398/sgbmD1fGWsVjFDBdGyp_gMw1ZxEWvbv-NoYbdYKC4_IZOtwGIH62TQQlwWrhh0PHFudZ"

def send_score_to_discord(score, username="Player"):
    # Create the message payload
    message = {
        "content": f"**Game Over!**\n**Score**: {score}\n**Player**: {username}",
        "embeds": [{
            "title": "Game Over",
            "description": f"{username} scored {score} points!",
            "color": 16711680  # Red color
        }]
    }
    
    # Send the POST request to the webhook URL
    try:
        response = requests.post(WEBHOOK_URL, json=message)
        if response.status_code == 204:
            print("Score sent to Discord successfully!")
        else:
            print(f"Failed to send score to Discord: {response.status_code}")
    except Exception as e:
        print(f"Error sending score to Discord: {e}")

def draw_grid():
    for x in range(0, WIDTH, CELL_SIZE):
        pygame.draw.line(screen, WHITE, (x, 0), (x, HEIGHT))
    for y in range(0, HEIGHT, CELL_SIZE):
        pygame.draw.line(screen, WHITE, (0, y), (WIDTH, y))

def place_food(snake, obstacles):
    while True:
        x = random.randint(0, WIDTH // CELL_SIZE - 1) * CELL_SIZE
        y = random.randint(0, HEIGHT // CELL_SIZE - 1) * CELL_SIZE
        if (x, y) not in snake and (x, y) not in obstacles:
            return x, y

def move_chaser(chaser, snake_head):
    cx, cy = chaser
    sx, sy = snake_head

    dx = sx - cx
    dy = sy - cy

    if abs(dx) > abs(dy):
        cx += CELL_SIZE if dx > 0 else -CELL_SIZE
    else:
        cy += CELL_SIZE if dy > 0 else -CELL_SIZE

    return cx, cy

def teleport_chaser(snake):
    while True:
        chaser_x = random.randint(0, (WIDTH // CELL_SIZE) - 1) * CELL_SIZE
        chaser_y = random.randint(0, (HEIGHT // CELL_SIZE) - 1) * CELL_SIZE
        if (chaser_x, chaser_y) not in snake:  # Check if the chaser is not on the snake
            return chaser_x, chaser_y

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

        # Render the current text.
        txt_surface = font.render(text, True, color)
        width = max(200, txt_surface.get_width()+10)
        input_box.w = width
        screen.blit(txt_surface, (input_box.x+5, input_box.y+5))
        pygame.draw.rect(screen, color, input_box, 2)

        # Display instructions
        prompt = font.render("Enter your username (press Enter to start):", True, WHITE)
        screen.blit(prompt, (WIDTH // 2 - prompt.get_width() // 2, HEIGHT // 4))

        pygame.display.flip()
        clock.tick(30)

def game_over(score, username):
    font = pygame.font.Font(None, 50)
    screen.fill(BLACK)
    text = font.render(f"Game Over! Score: {score}", True, WHITE)
    retry_text = font.render("Press R to Retry", True, WHITE)
    screen.blit(text, (WIDTH // 2 - text.get_width() // 2, HEIGHT // 3))
    screen.blit(retry_text, (WIDTH // 2 - retry_text.get_width() // 2, HEIGHT // 2))
    pygame.display.flip()

    # Send score to Discord when game is over
    send_score_to_discord(score, username)

    waiting_for_retry = True
    while waiting_for_retry:
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                pygame.quit()
                sys.exit()
            if event.type == pygame.KEYDOWN:
                if event.key == pygame.K_r:
                    waiting_for_retry = False  # Exit the loop to restart the game
                    return  # Retry game

def main():
    username = get_username()  # Prompt user for username at the start

    while True:
        # Snake starting position
        snake = [(WIDTH // 2, HEIGHT // 2)]
        direction = random.choice([UP, DOWN, LEFT, RIGHT])
        food = place_food(snake, [])
        obstacles = []

        # Chasers - starting with 3, and adding more as the score increases
        chasers = [teleport_chaser(snake) for _ in range(3)]  # Initial 3 chasers, none on the snake

        # Game speed (faster for harder gameplay)
        speed = 20

        # Score
        score = 0

        while True:
            screen.fill(BLACK)

            # Handle events
            for event in pygame.event.get():
                if event.type == pygame.QUIT:
                    pygame.quit()
                    sys.exit()

            # Controls
            keys = pygame.key.get_pressed()
            if keys[pygame.K_UP] and direction != DOWN:
                direction = UP
            if keys[pygame.K_DOWN] and direction != UP:
                direction = DOWN
            if keys[pygame.K_LEFT] and direction != RIGHT:
                direction = LEFT
            if keys[pygame.K_RIGHT] and direction != LEFT:
                direction = RIGHT

            # Move snake
            head_x, head_y = snake[0]
            new_head = (head_x + direction[0] * CELL_SIZE, head_y + direction[1] * CELL_SIZE)

            # Check for collisions
            if (new_head in snake or 
                new_head[0] < 0 or new_head[1] < 0 or 
                new_head[0] >= WIDTH or new_head[1] >= HEIGHT or 
                new_head in obstacles or 
                new_head in chasers):
                game_over(score, username)
                break  # Exit game loop and retry

            # Add new head to snake
            snake.insert(0, new_head)

            # Check if food is eaten
            if new_head == food:
                food = place_food(snake, obstacles)
                score += 1
                # Increase speed after every point
                speed += 1
                # Add more obstacles every few points
                if score % 3 == 0:
                    obstacles.append(place_food(snake, obstacles))
            else:
                snake.pop()  # Remove the tail if no food is eaten

            # Move the chasers faster based on score
            for _ in range(1 + score // 3):  # Increases chaser speed as score increases
                chasers = [move_chaser(chaser, snake[0]) for chaser in chasers]

            # Occasionally teleport a chaser to a random location to make it harder
            if random.random() < 0.1:  # 10% chance
                chasers[random.randint(0, len(chasers) - 1)] = teleport_chaser(snake)

            # Draw the food
            pygame.draw.rect(screen, RED, (food[0], food[1], CELL_SIZE, CELL_SIZE))

            # Draw the snake
            for segment in snake:
                pygame.draw.rect(screen, GREEN, (segment[0], segment[1], CELL_SIZE, CELL_SIZE))

            # Draw the obstacles
            for obstacle in obstacles:
                pygame.draw.rect(screen, BLUE, (obstacle[0], obstacle[1], CELL_SIZE, CELL_SIZE))

            # Draw the chasers
            for chaser in chasers:
                pygame.draw.rect(screen, YELLOW, (chaser[0], chaser[1], CELL_SIZE, CELL_SIZE))

            # Draw the grid
            draw_grid()

            # Update the display
            pygame.display.flip()

            # Control the game speed
            clock.tick(speed)

# Run the game
if __name__ == "__main__":
    main()
