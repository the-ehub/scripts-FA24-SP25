/**
 * K-Means Clustering Utility Functions
 *
 * This file implements a basic K-Means clustering algorithm for use in scenarios 
 * where you want to group vectors (e.g., numerical representations of student preferences, 
 * availability patterns, or interest embeddings).
 *
 * ✅ Usage:
 *   - Prepare your data as an array of vectors (each vector is an array of numbers).
 *   - Call `kMeansClustering(data, k)` to group the data into `k` clusters.
 *   - The result will be an array of objects: `{ vector: [...], cluster: number }`.
 *
 * ⚠️ Notes:
 *   - Centroids are initialized using the first k points (no randomization).
 *   - Euclidean distance is used to compute similarity between points.
 *   - There is no normalization, so input vectors should be scaled as needed beforehand.
 *   - The smaller helper functions (`initializeCentroids`, `assignCluster`, etc.) are internal 
 *     and typically do not need to be called directly.
 */

function kMeansClustering(data, k) {
  let maxIterations = 100;
  let centroids = initializeCentroids(data, k);
  let clusters = Array(data.length).fill(-1);
  
  for (let iter = 0; iter < maxIterations; iter++) {
    let newClusters = data.map(point => assignCluster(point, centroids));
    if (JSON.stringify(newClusters) === JSON.stringify(clusters)) break; // Stop if clusters don’t change
    clusters = newClusters;
    centroids = updateCentroids(data, clusters, k);
  }
  
  return data.map((point, index) => ({ vector: point, cluster: clusters[index] }));
}

function initializeCentroids(data, k) {
  return data.slice(0, k); // Pick first k points as initial centroids
}

function assignCluster(point, centroids) {
  let distances = centroids.map(centroid => euclideanDistance(point, centroid));
  return distances.indexOf(Math.min(...distances));
}

function updateCentroids(data, clusters, k) {
  let newCentroids = Array(k).fill(null).map(() => Array(data[0].length).fill(0));
  let counts = Array(k).fill(0);
  
  data.forEach((point, i) => {
    let cluster = clusters[i];
    counts[cluster]++;
    point.forEach((val, j) => newCentroids[cluster][j] += val);
  });
  
  return newCentroids.map((centroid, i) => centroid.map(val => (counts[i] ? val / counts[i] : 0)));
}

function euclideanDistance(point1, point2) {
  return Math.sqrt(point1.reduce((sum, val, i) => sum + Math.pow(val - point2[i], 2), 0));
}
