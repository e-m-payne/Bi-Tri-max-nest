import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt

# Function to create bipartite network from an Excel adjacency matrix
def create_bipartite_network_from_excel(file_path, sheet_name):
    # Read the Excel file into a DataFrame from the specified sheet
    df = pd.read_excel(file_path, sheet_name=sheet_name, index_col=0)

    # Create an empty bipartite graph
    B = nx.Graph()

    # Add plant nodes (rows) and pollinator nodes (columns) to the bipartite graph
    plants = df.index.tolist()  # Row names (plants)
    pollinators = df.columns.tolist()  # Column names (pollinators)

    B.add_nodes_from(plants, bipartite=0)  # Add plants as one set
    B.add_nodes_from(pollinators, bipartite=1)  # Add pollinators as another set

    # Add edges where the adjacency matrix has a '1'
    for plant in plants:
        for pollinator in pollinators:
            if df.loc[plant, pollinator] == 1:  # Check for interaction
                B.add_edge(plant, pollinator)

    return B, plants, pollinators

# Function to calculate the fraction of remaining plants
def calculate_remaining_plants(B, plants):
    remaining_plants = 0
    for plant in plants:
        if len(B[plant]) > 0:  # If plant still has connections
            remaining_plants += 1
    return remaining_plants / len(plants)

# Function to calculate robustness R
def calculate_robustness(B, plants, pollinators):
    robustness_curve = []

    while pollinators:
        # Find the pollinator with the largest degree (most connections)
        pollinator_degrees = [(pollinator, B.degree(pollinator)) for pollinator in pollinators]
        pollinator_to_remove = max(pollinator_degrees, key=lambda x: x[1])[0]

        # Remove that pollinator
        B.remove_node(pollinator_to_remove)
        pollinators.remove(pollinator_to_remove)

        # Calculate remaining fraction of plants
        remaining_plants_fraction = calculate_remaining_plants(B, plants)
        robustness_curve.append((len(pollinators) / len(plants), remaining_plants_fraction))

    # Calculate area under the curve
    area_under_curve = sum([r[1] for r in robustness_curve]) / len(robustness_curve)
    robustness_R = area_under_curve

    return robustness_curve, robustness_R

# Plot robustness curve (backwards)
def plot_robustness_curve(robustness_curve):
    # Reverse the x-axis to go from 1 to 0 (pollinators removed)
    x_vals = [1 - r[0] for r in robustness_curve]  # Reverse fraction of remaining pollinators
    y_vals = [r[1] for r in robustness_curve]

    plt.plot(x_vals, y_vals, marker='o')
    plt.xlabel("Fraction of pollinators remaining")
    plt.ylabel("Fraction of remaining plants")
    plt.title("Ecological Robustness of Bipartite Network (Targeted Removal)")
    plt.grid(True)
    plt.show()

# Main execution
if __name__ == "__main__":
    # Path to your Excel file
    file_path = '9bus bipartite matrices.xlsx'
    
    # Sheet name in the Excel file
    sheet_name = 'DS9'  # You can change this to the desired sheet name
    
    # Create bipartite network from the Excel file and specified sheet
    B, plants, pollinators = create_bipartite_network_from_excel(file_path, sheet_name)

    # Calculate robustness and plot curve
    robustness_curve, robustness_R = calculate_robustness(B, plants, pollinators)

    print(f"Robustness (R): {robustness_R:.3f}")

    # Plot the robustness curve (backwards)
    plot_robustness_curve(robustness_curve)


