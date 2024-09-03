# Filename: graph.py
# Author: Tyler Stringer
# Date: 2024-09-03
# Details: Displays formula/sheet dependencies as a graph network. Creates a directed graph where nodes represent
#   sheets and edges represent dependencies.

import networkx as nx
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.cm as cm
from networkx.drawing.nx_pydot import graphviz_layout


# Graph for showing sheet dependencies (use this if you need to run main.py for 1st time for this file). Run main.py:
def generate_graph(dependencies, layout):
    # Create a directed graph
    my_graph = nx.DiGraph()

    # Add nodes and edges
    for sheet, deps in dependencies.items():
        for dep in deps:
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
#        colors = cm.rainbow([i / len(edges) for i in range(len(edges))])  # Use a colormap
#        colors = cm.get_cmap('rainbow')(range(len(edges)))  # Use a colormap via get_cmap

    # Draw the edges (arrows) on graph
    draw_edges(my_graph, pos)

    # Draw labels
    nx.draw_networkx_labels(my_graph, pos, font_size=10, font_weight='bold')

    plt.title('Sheet Dependencies Overview')
    plt.show(block=False)  # Display this figure without blocking


'''
# Formula dependencies ... too many formulas to be a useful graphic right now:
def generate_formula_graph(formula_dependencies):
    # Create a directed graph for formulas
    graph = nx.DiGraph()

    # Add nodes and edges
    for formula, deps in formula_dependencies.items():
        for dep in deps:
            graph.add_edge(formula, dep)

    # Use pydot for hierarchical structure
    pos = graphviz_layout(graph, prog='dot')

    # Draw the graph
    plt.figure(figsize=(15, 10))

    # Draw nodes
    nx.draw_networkx_nodes(graph, pos, node_size=3000, node_color='lightgreen', alpha=0.5)

    # Draw edges with different colors
    edges = graph.edges()
    colors = [f"#{random.randint(0, 0xFFFFFF):06x}" for _ in range(len(edges))]  # Random colors for each arrow/edge
    nx.draw_networkx_edges(graph, pos, edgelist=edges, edge_color=colors, arrowstyle='-|>', arrowsize=20)

    # Draw labels
    nx.draw_networkx_labels(graph, pos, font_size=10, font_weight='bold')

    plt.title('Formula Dependencies Overview')
    plt.show(block=False)  # Display this figure without blocking
'''
