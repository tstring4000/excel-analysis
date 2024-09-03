# Filename: graph_existing.py
# Author: Tyler Stringer
# Date: 2024-09-03
# Details: Displays formula/sheet dependencies as a graph network. Creates a directed graph where nodes represent
#   sheets and edges represent dependencies. Run this only iff main.py has been run once for the Excel file.

import csv
import networkx as nx
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.cm as cm
import scipy as sp
from networkx.drawing.nx_pydot import graphviz_layout

# Filepath to existing dependencies csv:
dependencies_csv = 'data/output/dependencies.csv'


# Graph for showing sheet dependencies (use this if you already ran main.py once)
def load_dependencies(filename):
    # Loads the dependencies from a CSV file into a dictionary (must have run main.py at least once for the Excel file).
    # If you want to edit the dependencies, delete rows you no longer need from dependencies.csv:

    # Populate the dependencies dictionary automatically from file:
    existing_dependencies = {}
    with open(filename, mode='r') as file:
        reader = csv.reader(file)
        for row in reader:
            sheet = row[0].strip()
            deps = [dep.strip() for dep in row[1:] if dep.strip()]  # Filter out empty strings
            if sheet:  # Only add the sheet if it's not an empty string
                existing_dependencies[sheet] = deps
    return existing_dependencies


# Make existing_dependencies globally available after the script runs
existing_dependencies = load_dependencies(dependencies_csv)   # From csv file

# Print the dictionary to the console
print("Existing Dependencies Dictionary:")
print(existing_dependencies)


# Graph for showing sheet dependencies:
def generate_graph(existing_dependencies, layout):
    # Create a directed graph
    my_graph = nx.DiGraph()

    # Add nodes and edges
    for sheet, deps in existing_dependencies.items():
        if sheet:  # Ensure the sheet name is not empty
            for dep in deps:
                if dep:  # Ensure the dependency is not empty
                    my_graph.add_edge(sheet, dep)

    # Choose layout
    if layout == 'spring':
        pos = nx.spring_layout(my_graph)
    elif layout == 'circular':
        pos = nx.circular_layout(my_graph)
    elif layout == 'shell':
        pos = nx.shell_layout(my_graph)
    elif layout == 'planar':
        pos = nx.planar_layout(my_graph)
    elif layout == 'kamada_kawai':
        pos = nx.kamada_kawai_layout(my_graph)
    else:
        raise ValueError(
            "Invalid layout choice. Please choose from 'spring', 'circular', 'shell', 'planar', or 'kamada_kawai'.")

    # Draw the graph
    plt.figure(figsize=(10, 8))

    # Draw nodes
    nx.draw_networkx_nodes(my_graph, pos, node_size=3000, node_color='lightblue', alpha=0.5)

    # Draw edges with different colors
    def draw_edges(my_graph, pos):
        edges = my_graph.edges()
        n_edges = len(edges)    # Number of edges
        norm = np.linspace(0, 1, n_edges)   # Normalize edge indices to range [0, 1] so we get all colors
        cmap = cm.get_cmap('rainbow', n_edges)  # Get a colormap with n_edges amount of colors
        colors = cmap(norm)     # Apply the colormap to normalized values

        nx.draw_networkx_edges(my_graph, pos, edgelist=edges, edge_color=colors, arrowstyle='-|>', arrowsize=30,
                               connectionstyle='arc3,rad=0.1')

    # Draw the edges (arrows) on graph
    draw_edges(my_graph, pos)

    # Draw labels
    nx.draw_networkx_labels(my_graph, pos, font_size=10, font_weight='bold')

    plt.title('Existing Sheet Dependencies Overview')
    plt.show()


# If you want to run the graph generation directly:
if __name__ == "__main__":
    existing_dependencies = load_dependencies(dependencies_csv)

    # Layout: pick 'spring', 'circular', 'shell', 'planar', or 'kamada_kawai'
    generate_graph(existing_dependencies, 'circular')

    # Return the dictionary if the script is imported
    if 'existing_dependencies' not in globals():
        globals()['existing_dependencies'] = existing_dependencies
