import numpy as np
import matplotlib.pyplot as plt
import scipy.stats as stats
import pandas as pd

# Parameters for the log-normal distribution
mean = 200
variance = 20

# Calculate the shape (sigma) and scale (exp(mu)) parameters for the log-normal distribution
sigma = np.sqrt(np.log(1 + (variance / mean ** 2)))
mu = np.log(mean) - 0.5 * sigma ** 2

# Generate log-normal distribution data
data = np.random.lognormal(mean=mu, sigma=sigma, size=1000)

# Plot the histogram of the log-normal distribution
plt.figure(figsize=(10, 6))
plt.hist(data, bins=50, density=True, alpha=0.6, color='g')

# Plot the PDF of the log-normal distribution
xmin, xmax = plt.xlim()
x = np.linspace(xmin, xmax, 100)
p = (np.exp(-(np.log(x) - mu) ** 2 / (2 * sigma ** 2)) / (x * sigma * np.sqrt(2 * np.pi)))
plt.plot(x, p, 'k', linewidth=2)

title = "Log-Normal Distribution\nMean = 200, Variance = 20"
plt.title(title)
plt.xlabel('Value')
plt.ylabel('Frequency')
plt.show()

# Given parameters
mean = 509
sigma = 0.180227500468749
gamma = 0.5772

# Calculate the location (mu) and scale (beta) parameters for the Gumbel distribution
# z_80 = stats.gumbel_r.ppf(0.8)  # Z-score for the 80th percentile of the Gumbel distribution
# z_80 = stats.norm.ppf(0.8)
# print(z_80)
# std_dev = (fractile_80 - mean) / z_80
# std_dev = 0.3 * mean
# variance = std_dev ** 2
# print('Gumbel:', variance)
# beta = (fractile_80 - mean) / z_80
beta = mean * sigma * np.sqrt(6) / np.pi
mu = mean - beta * np.euler_gamma
# Calcul du fractile 80%
fractile_80 = stats.gumbel_r.ppf(0.8, loc=mu, scale=beta)
print('Gumbel fractile:', fractile_80)

# Generate Gumbel distribution data
data = np.random.gumbel(loc=mu, scale=beta, size=1000)

# Plot the histogram of the Gumbel distribution
plt.figure(figsize=(10, 6))
plt.hist(data, bins=50, density=True, alpha=0.6, color='r')

# Plot the PDF of the Gumbel distribution
xmin, xmax = plt.xlim()
x = np.linspace(xmin, xmax, 100)
p = stats.gumbel_r.pdf(x, loc=mu, scale=beta)
plt.plot(x, p, 'k', linewidth=2)

title = f"Gumbel Distribution\nMean = {mean}, Fractile 80% = {fractile_80}"
plt.title(title)
plt.xlabel('Value')
plt.ylabel('Density')
plt.xlim(0, 1500)
plt.xticks(np.arange(0, 1501, 100))
plt.show()

# Given parameters
mean = 509
fractile_80 = 575

# Calculate the standard deviation using the 80th percentile (fractile)
z_80 = stats.norm.ppf(0.8)  # Z-score for the 80th percentile
print(z_80)
std_dev = (fractile_80 - mean) / z_80
variance = std_dev ** 2
print('Normale:', variance)

# Generate normal distribution data
data = np.random.normal(loc=mean, scale=std_dev, size=1000)

# Plot the histogram of the normal distribution
plt.figure(figsize=(10, 6))
plt.hist(data, bins=50, density=True, alpha=0.6, color='b')

# Plot the PDF of the normal distribution
xmin, xmax = plt.xlim()
x = np.linspace(xmin, xmax, 100)
p = stats.norm.pdf(x, mean, std_dev)
plt.plot(x, p, 'k', linewidth=2)

title = f"Normal Distribution\nMean = {mean}, Fractile 80% = {fractile_80}, Variance = {variance}"
plt.title(title)
plt.xlabel('Value')
plt.ylabel('Density')
plt.xlim(0, 1500)
plt.xticks(np.arange(0, 1501, 100))
# plt.show()
# plt.savefig('D:\\Donnees\\Thèse\\Calculs Feu\\Distributions\\normal.png')
plt.show()

# # Given parameters
# mean = 780
# fractile_80 = 948
#
# # Calculate the Z-score for the 80th percentile
# z_80 = stats.norm.ppf(0.8)
#
# # Calculate the standard deviation (sigma)
# std_dev = (fractile_80 - mean) / z_80
#
# # Calculate the variance (sigma^2)
# variance = std_dev ** 2
# print('Normale:', variance)

# Parameters for the normal distribution
mean = 200
variance = 20
std_dev = np.sqrt(variance)
lower_bound = 195
upper_bound = 500

# Generate normal distribution data within the specified bounds
data = np.random.normal(loc=mean, scale=std_dev, size=1000)
data = data[(data >= lower_bound) & (data <= upper_bound)]

# Plot the histogram of the normal distribution
plt.figure(figsize=(10, 6))
plt.hist(data, bins=50, density=True, alpha=0.6, color='b')

# Plot the PDF of the normal distribution
xmin, xmax = plt.xlim()
x = np.linspace(xmin, xmax, 100)
p = stats.norm.pdf(x, mean, std_dev)
plt.plot(x, p, 'k', linewidth=2)

title = (f"Normal Distribution\nMean = {mean}, Variance = {variance}, Lower Bound = {lower_bound}, "
         f"Upper Bound = {upper_bound}")
plt.title(title)
plt.xlabel('Value')
plt.ylabel('Frequency')
plt.show()

# Given parameters
mean = 509
sigma = 0.180227500468749

# Calculate the location (mu) and scale (beta) parameters for the Gumbel distribution
# z_80 = stats.gumbel_r.ppf(0.8)  # Z-score for the 80th percentile of the Gumbel distribution
# beta = (fractile_80 - mean) / z_80
beta = mean * sigma * np.sqrt(6) / np.pi
mu = mean - beta * np.euler_gamma
size = 1000

# Generate Gumbel distribution data
data = np.random.gumbel(loc=mu, scale=beta, size=size)

# Plot the histogram of the Gumbel distribution
fig, ax1 = plt.subplots(figsize=(10, 6))

color = 'tab:red'
ax1.set_xlabel('Value')
ax1.set_ylabel('Density', color=color)
ax1.hist(data, bins=50, density=True, alpha=0.6, color=color)
ax1.tick_params(axis='y', labelcolor=color)

# Plot the PDF of the Gumbel distribution
xmin, xmax = ax1.get_xlim()
x = np.linspace(xmin, xmax, 100)
p = stats.gumbel_r.pdf(x, loc=mu, scale=beta)
ax1.plot(x, p, 'k', linewidth=2)

ax2 = ax1.twinx()  # instantiate a second axes that shares the same x-axis

color = 'tab:blue'
ax2.set_ylabel('Number of Draws', color=color)  # we already handled the x-label with ax1
counts, bins = np.histogram(data, bins=50)
ax2.plot(bins[:-1], counts, 'o', color=color)
ax2.tick_params(axis='y', labelcolor=color)

fig.tight_layout()  # otherwise the right y-label is slightly clipped
plt.title(f"Gumbel Distribution\nMean = {mean}, Fractile 80% = {fractile_80}")
plt.xlim(0, 1500)
plt.xticks(np.arange(0, 1501, 100))

# Save the figure
plt.savefig('gumbel_distribution_with_counts.png')
plt.show()
# Create a DataFrame
df = pd.DataFrame(data, columns=[f"Gumbel Distribution\nMean = {mean}, Fractile 80% = {fractile_80}"])

# Save the DataFrame to an Excel file
# CstbGroup = r'C:\Users\francois.consigny\CSTBGroup\DSSF_Projet_These_Francois_Consigny - Documents'
# cur_path = (r'C:\Users\Francois.CONSIGNY\CSTBGroup\DSSF_Projet_These_Francois_Consigny - Documents'
#             r'\Calculs Feu\Distributions')
cur_path = (r'C:\Users\franc\CSTBGroup\These_Francois_Consigny - Documents'
            r'\Calculs Feu\Distributions')
df.to_excel(cur_path + f"\\Gumbel_distribution_with_lower_bound_{size}.xlsx", index=False)

print("Gumbel distribution data  has been saved to "
      f"\\Gumbel_distribution_with_lower_bound_{size}.xlsx")

# Opening factor (Log Normal lower bound 0.06)
# Parameters for the log-normal distribution
mean = 0.07
variance = 0.001

# Calculate the shape (sigma) and scale (exp(mu)) parameters for the log-normal distribution
sigma = np.sqrt(np.log(1 + (variance / mean ** 2)))
mu = np.log(mean) - 0.5 * sigma ** 2

# Generate log-normal distribution data with a lower bound of 0.06
data = []
while len(data) < 1000:
    value = np.random.lognormal(mean=mu, sigma=sigma)
    if value >= 0.06:
        data.append(value)

# Create a DataFrame
df = pd.DataFrame(data, columns=['LogNormalDistribution'])

# Save the DataFrame to an Excel file
# CstbGroup = r'C:\Users\francois.consigny\CSTBGroup\DSSF_Projet_These_Francois_Consigny - Documents'
# cur_path = (r'C:\Users\Francois.CONSIGNY\CSTBGroup\DSSF_Projet_These_Francois_Consigny - Documents'
#             r'\Calculs Feu\Distributions')
cur_path = (r'C:\Users\franc\CSTBGroup\These_Francois_Consigny - Documents'
            r'\Calculs Feu\Distributions')
df.to_excel(cur_path + r'\log_normal_distribution_with_lower_bound.xlsx', index=False)

print("Log-normal distribution data with a lower bound has been saved to "
      "'log_normal_distribution_with_lower_bound.xlsx'.")

# Opening factor (Log Normal lower bound 0.06)
# Parameters for the log-normal distribution
mean = 0.07
variance = 0.001

# Calculate the shape (sigma) and scale (exp(mu)) parameters for the log-normal distribution
sigma = np.sqrt(np.log(1 + (variance / mean ** 2)))
mu = np.log(mean) - 0.5 * sigma ** 2

# Generate log-normal distribution data with a lower bound of 0.06
data = []
while len(data) < 1000:
    value = np.random.lognormal(mean=mu, sigma=sigma)
    if value >= 0.06:
        data.append(value)

# Create a DataFrame
df = pd.DataFrame(data, columns=['LogNormalDistribution'])

# Save the DataFrame to an Excel file
# CstbGroup = r'C:\Users\francois.consigny\CSTBGroup\DSSF_Projet_These_Francois_Consigny - Documents'
# cur_path = (r'C:\Users\Francois.CONSIGNY\CSTBGroup\DSSF_Projet_These_Francois_Consigny - Documents'
#             r'\Calculs Feu\Distributions')
cur_path = (r'C:\Users\franc\CSTBGroup\These_Francois_Consigny - Documents'
            r'\Calculs Feu\Distributions')
df.to_excel(cur_path + r'\log_normal_distribution_with_lower_bound.xlsx', index=False)

print("Log-normal distribution data with a lower bound has been saved to "
      "'log_normal_distribution_with_lower_bound.xlsx'.")
