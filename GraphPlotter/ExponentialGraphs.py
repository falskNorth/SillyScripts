import numpy as np
import matplotlib.pyplot as plt

# First dataset (Quadratic Fit)
x = np.array([6645, 8389, 9648, 11420, 13367, 15816, 18139, 20907, 23988, 28146, 30124])  # X values // Voltage (Siemens)
y = np.array([50, 45, 40, 35, 30, 25, 20, 15, 10, 5, 0])  # Y values // Temperature

# Perform polynomial regression (degree 2 for a quadratic fit)
coefficients = np.polyfit(x, y, 2)  # 2 indicates quadratic fit
quadratic_fit = np.poly1d(coefficients)

# Generate x values for the curve of best fit
x_fit = np.linspace(min(x), max(x), 100)
y_fit = quadratic_fit(x_fit)

# Plotting the arrays against each other
plt.figure(figsize=(10, 5))

# SCATTERPLOT
plt.scatter(x, y, color='blue', label='Data Points', marker='o')

# CURVE PLOT for the curve of best fit
plt.plot(x_fit, y_fit, color='red', label='Curve of Best Fit')

# ADDING LABELS
plt.title('Temperature vs Volts through PLC (CURVED)')
plt.xlabel('Voltage (Siemens)')
plt.ylabel('Temperature')
plt.legend()  # SHOW LEGEND
plt.grid()  # SHOW GRID

# Display the equation of the curve
equation_text = f'y = {coefficients[0]:.2f}x² + {coefficients[1]:.2f}x + {coefficients[2]:.2f}'
plt.text(0.05, 0.95, equation_text, transform=plt.gca().transAxes, fontsize=12, verticalalignment='top')
print("Quadratic Fit Equation:", equation_text)

# Show the plot
plt.show()

# Second dataset (Linear Fit)
x = np.array([7380, 8463, 10140, 11630, 13700, 16023, 18592, 21465, 24850, 27084, 30512])  # X values // Voltage (Siemens)
y = np.array([50, 45, 40, 35, 30, 25, 20, 15, 10, 5, 0])  # Y values // Temperature

# Perform polynomial regression (degree 1 for a linear fit)
coefficients = np.polyfit(x, y, 1)  # 1 indicates linear fit
linear_fit = np.poly1d(coefficients)

# Generate x values for the curve of best fit
x_fit = np.linspace(min(x), max(x), 100)
y_fit = linear_fit(x_fit)

# Plotting the arrays against each other
plt.figure(figsize=(10, 5))

# SCATTERPLOT
plt.scatter(x, y, color='blue', label='Data Points', marker='o')

# CURVE PLOT for the curve of best fit
plt.plot(x_fit, y_fit, color='red', label='Curve of Best Fit')

# ADDING LABELS
plt.title('Temperature vs Volts through PLC (LINEAR)')
plt.xlabel('Voltage (Siemens)')
plt.ylabel('Temperature')
plt.legend()  # SHOW LEGEND
plt.grid()  # SHOW GRID

# Display the equation of the curve
equation_text = f'y = {coefficients[0]:.2f}x + {coefficients[1]:.2f}'
plt.text(0.05, 0.95, equation_text, transform=plt.gca().transAxes, fontsize=12, verticalalignment='top')
print("Linear Fit Equation:", equation_text)

# Show the plot
plt.show()