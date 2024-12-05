let canvas = document.getElementById('gameCanvas');
let ctx = canvas.getContext('2d');

let cellSize = 20;
let width = canvas.width;
let height = canvas.height;
let score = 0;
let snake = [{x: width / 2, y: height / 2}];
let direction = 'RIGHT';
let food = generateFood();
let username = '';

let chasers = [];
let speed = 100;  // Initial speed

// Directions
const DIRECTIONS = {
    UP: {x: 0, y: -cellSize},
    DOWN: {x: 0, y: cellSize},
    LEFT: {x: -cellSize, y: 0},
    RIGHT: {x: cellSize, y: 0},
};

function startGame() {
    username = document.getElementById('usernameInput').value || 'Player';
    document.getElementById('username').style.display = 'none'; // Hide the username input form
    gameLoop(); // Start the game loop
    document.addEventListener('keydown', changeDirection); // Listen for keydown events to change direction
}

function changeDirection(event) {
    if (event.key === 'ArrowUp' && direction !== 'DOWN') {
        direction = 'UP';
    } else if (event.key === 'ArrowDown' && direction !== 'UP') {
        direction = 'DOWN';
    } else if (event.key === 'ArrowLeft' && direction !== 'RIGHT') {
        direction = 'LEFT';
    } else if (event.key === 'ArrowRight' && direction !== 'LEFT') {
        direction = 'RIGHT';
    }
}

function generateFood() {
    let foodX = Math.floor(Math.random() * (width / cellSize)) * cellSize;
    let foodY = Math.floor(Math.random() * (height / cellSize)) * cellSize;
    return {x: foodX, y: foodY};
}

function gameLoop() {
    ctx.clearRect(0, 0, width, height); // Clear the canvas before drawing

    // Move snake
    let head = {x: snake[0].x + DIRECTIONS[direction].x, y: snake[0].y + DIRECTIONS[direction].y};
    snake.unshift(head); // Add new head to the snake

    // Check for collisions with walls or snake itself
    if (head.x < 0 || head.x >= width || head.y < 0 || head.y >= height || isCollisionWithSnake(head)) {
        gameOver();
        return;
    }

    // Check if snake eats food
    if (head.x === food.x && head.y === food.y) {
        score++;
        food = generateFood(); // Generate new food
        speed = Math.max(50, speed - 5);  // Increase speed
    } else {
        snake.pop(); // Remove the tail if no food is eaten
    }

    // Draw everything
    drawSnake();
    drawFood();
    drawScore();
    
    setTimeout(gameLoop, speed); // Recursive call to keep the game running
}

function drawSnake() {
    snake.forEach(segment => {
        ctx.fillStyle = 'green';
        ctx.fillRect(segment.x, segment.y, cellSize, cellSize);
    });
}

function drawFood() {
    ctx.fillStyle = 'red';
    ctx.fillRect(food.x, food.y, cellSize, cellSize);
}

function drawScore() {
    ctx.fillStyle = 'white';
    ctx.font = '20px Arial';
    ctx.fillText('Score: ' + score, 10, 30);
}

function isCollisionWithSnake(head) {
    for (let i = 1; i < snake.length; i++) {
        if (snake[i].x === head.x && snake[i].y === head.y) {
            return true; // Collision with snake body
        }
    }
    return false;
}

function gameOver() {
    ctx.clearRect(0, 0, width, height);
    ctx.fillStyle = 'white';
    ctx.font = '30px Arial';
    ctx.fillText('Game Over!', width / 2 - 90, height / 2);
    ctx.fillText('Score: ' + score, width / 2 - 60, height / 2 + 40);
    setTimeout(() => {
        document.getElementById('username').style.display = 'block'; // Show username input again
    }, 2000);
}

// Initialize game start screen
document.getElementById('username').style.display = 'block'; // Show username input initially
