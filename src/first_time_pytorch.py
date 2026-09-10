# https://www.youtube.com/watch?v=r1bquDz5GGA&t=898s

import torch
from tqdm import tqdm

# batch of data will have 10 points
N = 10

# each data point has one input and one output
D_in  = 1
D_out = 1

# create input data X
X = torch.randn(N, D_in)

# create output data y
y = torch.randn(N, D_out)

W_true = torch.tensor([[2.0]])
b_true = torch.tensor([[1.0]])
y_true = X @ W_true + b_true + torch.randn(N, D_out) * 0.1 # adding noise

# initialise teh paramters W and b with random values
W = torch.randn(D_in, D_out, requires_grad=True)
b = torch.randn(1, requires_grad=True)

print(f"Initial weight W: \n {W}\n")
print(f"Initial bias b: \n {b}\n")

# forward pass through the model
y_hat = X @ W + b

print(f"Predicted output y_pred (first three rows): \n {y_hat[:3]}\n")
print(f"True output y_true (first three rows): \n {y_hat[:3]}\n")

# calculate the loss
loss = torch.mean((y_hat - y_true) ** 2)
print(f"Loss: {loss.item():.4f}")

# backward pass to compute gradients
loss.backward() # tell Pytroch to travel backward from 'loss' and caclulate gradients for all parameters with 'requires_grad = True'

print(f"Gradient for W (dLoss/dW): \n {W.grad.item():.4f}\n")
print(f"Gradient for b (dLoss/db): \n {b.grad.item():.4f}")

# hyperparameters
learning_rate, epochs = 0.01, 1000
# an epoch is a single pass through the entire dataset

# re-initialise the parameters
W, b = torch.randn(D_in, D_out, requires_grad=True), torch.randn(1, requires_grad=True)

# training loop
for epoch in tqdm(range(epochs)):
    
    # forward pass and loss
    y_hat = X @ W + b
    loss = torch.mean((y_hat - y_true) ** 2)   
    
    # backward pass
    loss.backward()
    
    # update parameters using gradient descent
    with torch.no_grad():
        W -= learning_rate * W.grad
        b -= learning_rate * b.grad
        
        # zero the gradients after updating
        W.grad.zero_()
        b.grad.zero_()
    
    if epoch % 200 == 0:
        print(f"Epoch {epoch:02d}: Loss = {loss.item():.4f}, W = {W.item():.4f}, b = {b.item():.4f}")
        
print(f"\nFinal Parameters: W = {W.item():.4f}, b = {b.item():.4f}")
print(f"True parameters: W = 2.0000, b = 1.0000")
print(f"Learning rate of {learning_rate} and {epochs:,.0f} epochs")