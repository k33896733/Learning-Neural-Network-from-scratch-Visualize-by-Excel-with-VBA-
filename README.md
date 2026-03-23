## Background of the project

This project originated from my interest in understanding the inner workings of neural network (NN) without relying on high-level frameworks like TensorFlow or PyTorch. I chose VBA within Excel for the following reasons:

- **Accessibility** — Excel is widely available, requiring no additional installation for basic use.  (And I can study it during my office time)
-  **Neuron-level visibility** — Each cell in sheet "Forward" and "Backward" can directly represent an individual neuron or intermediate value, offering unparalleled educational clarity.

**The Model**
The model is trained on historical stock data (OHLCV and derived features) to predict the day closing price (change : close - open) (single-value regression).  
- Input: 10 features  (From HSI index)
- Output: 1 value  
- Hidden layers: 3 layers with Leaky ReLU activation (configurable slope α)  
- Training: Mini-batch gradient descent with L2 weight decay and batch normalization
**Note**: Indeed, I have tried different combination , but the performances are not good, so I just randomly pick one for illustration purpose

**Performance**

Financial time series are highly noisy and non-stationary. This implementation is **not** intended to deliver profitable trading signals and does **not** achieve competitive predictive accuracy compared to state-of-the-art models (e.g., LSTMs or industrial-scale networks).  

<img width="500" height="200" alt="image" src="https://github.com/user-attachments/assets/bbfccd92-9a95-4652-80ae-0c878ff468c8" />
<img width="500" height="200" alt="prediction2" src="https://github.com/user-attachments/assets/17364d62-7225-41f2-8d12-7943362addd2" />

**How to know if the model is correct?**
The primary goal is educational: to verify correct implementation of core neural network mechanics.  
A key validation test is that when **one of the input features is set to the target closing price itself**, the model rapidly learns to reproduce that value exactly — demonstrating that **the forward and backward propagation logic functions correctly**.

<img width="500" height="200" alt="image" src="https://github.com/user-attachments/assets/8630de3f-a14d-4521-b235-9cb5706b0610" />
<img width="500" height="200" alt="prediction" src="https://github.com/user-attachments/assets/b68b7cac-23fc-439b-92df-c84435a9aec8" />

## Neural Network

Neural networks are computational models inspired by biological neural systems. They learn to map inputs to desired outputs by iteratively adjusting internal parameters (weights and biases) through training.

<img width="300" height="300" alt="image" src="https://github.com/user-attachments/assets/c52af275-b87a-4c68-a998-6c426edfe3c5" />

### Basic Unit — Neuron

The fundamental building block of a neural network is the **neuron** (also called a node or artificial neuron). Each neuron performs the following operations:

1. Receives multiple input signals.
2. Computes a weighted sum of the inputs plus a bias term.
3. Applies a nonlinear activation function to the result.
4. Produces an output that is passed to the next layer (or serves as the final prediction).

### Forward Propagation

Forward propagation (or forward pass) is the process of computing the network's output given an input. Starting from the input layer, each subsequent layer computes its activations using the weights, biases, and activation functions from the previous layer's outputs. This continues until the output layer produces the final prediction.

<img width="500" height="200" alt="image" src="https://github.com/user-attachments/assets/98a719e4-b9ed-4bd9-8ce5-b3786e37161a" />

### Backward Propagation (Backpropagation)

To enable learning, the network minimizes a loss function that quantifies the difference between predicted and actual outputs (e.g., Mean Squared Error for regression). Backpropagation computes the gradients of the loss with respect to all weights and biases using the chain rule of calculus. These gradients indicate how each parameter should be adjusted to reduce the loss.

<img width="500" height="200" alt="image" src="https://github.com/user-attachments/assets/10073ce2-6d85-418f-8f82-d446675d526e" />

## How to Use This Excel Workbook

1. **Prerequisites**  
   - Microsoft Excel (2010 or later, with macros enabled).  
   - Enable macros when opening the file (File → Options → Trust Center → Macro Settings).

2. **Workbook Sheet Overview**  
   - **Data Raw** — Contains original and preprocessed historical stock data.
   - **Panel** — Prepares the 10 input features for each training example; includes the "Auto-learn" button to start training.
   - **Forward** — Visualizes node activations and computes forward propagation; adjustable parameters include Learning Rate and Leaky ReLU α.
   - **Backward** — Visualizes intermediate computations and performs backward propagation.
   - **Analysis** — (Template) Displays training progress, loss history, predictions vs. actual values, and evaluation metrics.
   - **Epoch(X)** — Dynamically created sheets recording loss and metrics for each training run.
   - **Note**: All core calculations occur in the **Forward** and **Backward** sheets via cell formulas. VBA primarily handles data copying, loop control, and sheet management.

3. **Basic Workflow**
   - Load and preprocess stock data in **Data Raw**.
   - In **Panel**, click **Update Data Set** to copy data to input nodes (or manually adjust inputs; unused nodes may be left blank).
   - Adjust hyperparameters in **Panel** and **Forward** sheets.
   - Click **Auto-learn** in **Panel** to start training.
   - Monitor progress in newly created **Epoch(X)** sheets and **Analysis**.
   - To restart: Click **Reset Model** in **Panel** and delete existing **Epoch(X)** sheets.

### Changing Input Data
- Input features are flexible: leave cells blank if fewer than 10 are needed.
- Use **Update Data Set** button for automatic refresh from **Data Raw**, or edit manually.
- The button executes simple VBA copy operations; the code is editable for customization.

**Note**: 
1. Due to VBA's loop-based nature and Excel's single-threaded execution, training on large datasets may take minutes to hours depending on hardware.
2. The formulas and VBA code are designed to be self-explanatory for readers familiar with neural network mathematics. Detailed cell/VBA logic is not documented line-by-line here.

## Contact / Feedback

If you find this project useful or have suggestions, please open an issue or contact me.
Thanks for checking out this implementation for documenting my learning.
