import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt

# Function to create bipartite network from an Excel adjacency matrix for a given sheet
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

# Function to calculate robustness R and compile elimination results into DataFrame
def calculate_robustness(B, plants, pollinators, sheet_name):
    robustness_curve = []
    results = []  # List to store results for the DataFrame

    while pollinators:
        # Find the pollinator with the largest degree (most connections)
        pollinator_degrees = [(pollinator, B.degree(pollinator)) for pollinator in pollinators]
        pollinator_to_remove = max(pollinator_degrees, key=lambda x: x[1])[0]

        # Remove that pollinator
        B.remove_node(pollinator_to_remove)
        pollinators.remove(pollinator_to_remove)

        # Calculate remaining fraction of plants
        remaining_plants_fraction = calculate_remaining_plants(B, plants)
        remaining_pollinators_fraction = len(pollinators) / len(plants)

        robustness_curve.append((remaining_pollinators_fraction, remaining_plants_fraction))

        # Store the results for this elimination step
        results.append({
            'Sheet': sheet_name,
            'Pollinators_Remaining_Fraction': remaining_pollinators_fraction,
            'Plants_Remaining_Fraction': remaining_plants_fraction
        })

    # Calculate area under the curve for robustness
    area_under_curve = sum([r[1] for r in robustness_curve]) / len(robustness_curve)
    robustness_R = area_under_curve

    return robustness_curve, robustness_R, pd.DataFrame(results)

# Plot robustness curve (backwards)
def plot_robustness_curve(robustness_curve, sheet_name):
    # Reverse the x-axis to go from 1 to 0 (pollinators removed)
    x_vals = [1 - r[0] for r in robustness_curve]  # Reverse fraction of remaining pollinators
    y_vals = [r[1] for r in robustness_curve]

    plt.plot(x_vals, y_vals, marker='o')
    plt.xlabel("Fraction of pollinators remaining")
    plt.ylabel("Fraction of remaining plants")
    plt.title(f"Ecological Robustness - {sheet_name}")
    plt.grid(True)
    plt.show()

# Main execution for processing multiple sheets and exporting results
if __name__ == "__main__":
    # Path to your Excel file
    file_path = 'Copy of 9bus_normal_optim_lim_matrix_step_145.xlsx'
    
    # List of sheet names to process
    sheet_names = ['normal', 'ds1', 'ds2', 'ds3', 'ds5', 'ds7', 'ds8', 'ds9']  # Add more sheet names as needed

    all_results = pd.DataFrame()  # Initialize empty DataFrame to store all results

    for sheet in sheet_names:
        # Create bipartite network from the Excel file for the current sheet
        B, plants, pollinators = create_bipartite_network_from_excel(file_path, sheet)

        # Calculate robustness and get elimination results
        robustness_curve, robustness_R, sheet_results = calculate_robustness(B, plants, pollinators, sheet)

        print(f"Robustness (R) for {sheet}: {robustness_R:.3f}")

        # Plot the robustness curve (optional)
        plot_robustness_curve(robustness_curve, sheet)

        # Append the results for the current sheet to the main DataFrame
        all_results = pd.concat([all_results, sheet_results], ignore_index=True)

    # Export all results to a new Excel file
    output_file = '9bus_NAPS_robustness_elimination_results.xlsx'
    all_results.to_excel(output_file, index=False)

    print(f"All elimination results have been saved to {output_file}")
