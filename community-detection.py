import json
import networkx as nx
import matplotlib.pyplot as plt
import seaborn as sns
import community as community_louvain  # pip install python-louvain
from collections import Counter, defaultdict
import pandas as pd
import numpy as np

"""
Community Detection via Interest Co-Occurrence Graphs

This script performs community detection among students based on overlapping interests.
It supports exploratory clustering, visual analytics, and group-level assignment for downstream applications like scheduling or team formation.

Main Functional Areas:

1. **Data Loading**:
   - Loads preprocessed student interest and availability data from a JSON file (`student_data.json`).
   - Loads a subset of target students (e.g., those in a particular pool or track) from a CSV file.

2. **Interest Graph Construction**:
   - Builds a co-occurrence graph of interests (nodes = interests, edges = shared selection).
   - Edges are weighted based on how often two interests co-occurred among students.
   - Constructs a co-occurrence matrix (DataFrame) for visual representation.

3. **Community Detection (Louvain Algorithm)**:
   - Applies Louvain clustering on the interest graph to detect natural communities of related interests.
   - Groups interests into clusters representing conceptual themes or affinity areas.

4. **Visualization**:
   - Displays the interest graph using networkx with nodes colored by cluster assignment.
   - Displays a heatmap of the interest co-occurrence matrix using seaborn.

5. **Student-to-Cluster Assignment**:
   - For each student, determines their strongest matching cluster based on shared interests.
   - Records the student's info, cluster assignment, and matched interests.

6. **Reporting & Export**:
   - Summarizes number of students per cluster and per track.
   - Shows the top shared interests within each cluster.
   - Outputs comma-separated email lists for potential calendar invites.
   - Saves all assignments to `student_interest_clusters.csv`.

Requirements:
- pandas
- numpy
- networkx
- seaborn
- matplotlib
- community (install via `pip install python-louvain`)

This script is designed to work alongside a Google Apps Script project that gathers and stores `student_data.json` and interest pool definitions.
"""

# Load the student data
with open('student_data.json', 'r') as f:
    student_data = json.load(f)

# Load target students from CSV
pool_df = pd.read_csv("micro-community-pool.csv")
pool_df['Email'] = pool_df['Email'].str.strip().str.lower()

# Define relevant interest categories (excluding ambiguous ones and AI)
excluded_interests = {"AI & machine learning", "something not listed", "still figuring it out"}
interest_pool = set()
for email in pool_df['Email']:
    student = student_data.get(email)
    if not student: continue
    interest_pool.update([i for i in student['interests'] if i not in excluded_interests])
interest_pool = sorted(interest_pool)

# Build co-occurrence graph and matrix
G = nx.Graph()
G.add_nodes_from(interest_pool)
co_matrix = pd.DataFrame(0, index=interest_pool, columns=interest_pool)

for email in pool_df['Email']:
    student = student_data.get(email)
    if not student: continue
    interests = [i for i in student['interests'] if i in interest_pool]
    for i in range(len(interests)):
        for j in range(i + 1, len(interests)):
            a, b = interests[i], interests[j]
            co_matrix.at[a, b] += 1
            co_matrix.at[b, a] += 1
            if G.has_edge(a, b):
                G[a][b]['weight'] += 1
            else:
                G.add_edge(a, b, weight=1)

# Apply Louvain community detection on interest graph
partition = community_louvain.best_partition(G, weight='weight', resolution=1.0)

# Organize interests by their cluster
interest_clusters = defaultdict(list)
for interest, cluster_id in partition.items():
    interest_clusters[cluster_id].append(interest)

# Visualize the graph with colored communities
pos = nx.spring_layout(G, seed=42)
colors = [partition[n] for n in G.nodes()]
plt.figure(figsize=(12, 8))
nx.draw_networkx(G, pos, node_color=colors, with_labels=True, node_size=800, font_size=10, cmap=plt.cm.tab10)
plt.title("Interest Co-Occurrence Graph with Louvain Clusters")
plt.axis('off')
plt.tight_layout()
plt.show()

# Visualize co-occurrence matrix as heatmap
plt.figure(figsize=(16, 14))
sns.heatmap(co_matrix, cmap='Reds', square=True, linewidths=.5, cbar_kws={'shrink': 0.5})
plt.title("Interest Co-Occurrence Heatmap", fontsize=16)
plt.xticks(rotation=90)
plt.yticks(rotation=0)
plt.tight_layout()
plt.show()

# Assign students to clusters within each track
student_assignments = []
for track in pool_df['Track'].unique():
    track_df = pool_df[pool_df['Track'] == track]
    for _, row in track_df.iterrows():
        email = row['Email']
        student = student_data.get(email)
        if not student: continue
        student_interests = set(i for i in student['interests'] if i in interest_pool)
        if not student_interests:
            continue
        cluster_scores = {
            cluster_id: len(student_interests & set(interest_list))
            for cluster_id, interest_list in interest_clusters.items()
        }
        top_cluster = max(cluster_scores, key=cluster_scores.get)
        student_assignments.append({
            'email': email,
            'firstName': student['firstName'],
            'lastName': student['lastName'],
            'track': track,
            'cluster': top_cluster,
            'matched_interests': list(student_interests & set(interest_clusters[top_cluster]))
        })

# View summary of clusters
assignment_df = pd.DataFrame(student_assignments)
print("\nCluster Summary (students per track and cluster):")
print(assignment_df.groupby(['track', 'cluster'])['email'].count())

# Show top interests in each cluster
print("\nTop Interests in Each Cluster:")
for cluster_id, interests in interest_clusters.items():
    print(f"\nCluster {cluster_id}:")
    counts = Counter()
    for email in assignment_df[assignment_df['cluster'] == cluster_id]['email']:
        student = student_data[email]
        counts.update(i for i in student['interests'] if i in interests)
    top_interests = counts.most_common(5)
    for interest, count in top_interests:
        print(f"  {interest} ({count})")

# Export comma-separated emails for each track/cluster combo
print("\nEmail lists for calendar invites:")
grouped_emails = assignment_df.groupby(['track', 'cluster'])['email'].apply(lambda x: ', '.join(x)).reset_index()
for _, row in grouped_emails.iterrows():
    print(f"Track: {row['track']} | Cluster: {row['cluster']}\nEmails: {row['email']}\n")

# Optional: save assignments to CSV
assignment_df.to_csv("student_interest_clusters.csv", index=False)
