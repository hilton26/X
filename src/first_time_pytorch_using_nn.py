# https://www.youtube.com/watch?v=r1bquDz5GGA&t=898s

import torch
from tqdm import tqdm

# batch of data will have 10 points
N = 10

# each data point has one input and one output; input has one fauture, output has one value
D_in  = 1
D_out = 1

# create input data X
X = torch.randn(N, D_in)

# create the linear layer
linear_layer = torch.nn.Linear(in_features = D_in, out_features = D_out)

# look inside the layer and see the parameters created
print(f"Layer's weight W: \n {linear_layer.weight}\n")
print(f"Initial bias b: \n {linear_layer.bias}\n")

# use it just like a function. This one is the forward pass
# (assume X is a tensor fo shape [10, 1])
y_hat = linear_layer(X)

print(f"Output of linear layer (first three rows): \n {y_hat[:3]}\n")

relu = torch.nn.ReLU()
sample_data = torch.tensor([-2.0, -.5, 0, .5, 2.0])
activated_data = relu(sample_data)
print(f"ReLU activation on sample data: \n {activated_data}\n")