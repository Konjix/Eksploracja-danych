import pandas as pd
import numpy as np
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import silhouette_score
from sklearn.cluster import KMeans
from scipy.cluster.hierarchy import dendrogram, linkage, fcluster
import matplotlib.pyplot as plt
import seaborn as sns



def find_clusters_elbow(data_tablica, show = True):
    # Usunięcie pierwszej kolumny zawierającej nazwy regionów
    clustering_data = data_tablica.drop(columns=data_tablica.columns[0])

    # Zastąpienie wszystkich '-' wartością NaN i konwersja wszystkich kolumn na typ float
    clustering_data = clustering_data.replace('-', np.nan).astype(float)

    # Wypełnienie brakujących wartości średnią dla każdej kolumny
    clustering_data = clustering_data.apply(lambda x: x.fillna(x.mean()), axis=0)

    # Normalizacja danych
    scaler = StandardScaler()
    scaled_data = scaler.fit_transform(clustering_data)

    # Obliczanie sumy kwadratów wewnętrznych (WSS) dla różnej liczby klastrów
    wss = []
    for i in range(1, 21):
        kmeans = KMeans(n_clusters=i, init='k-means++', max_iter=300, n_init=10, random_state=0)
        kmeans.fit(scaled_data)
        wss.append(kmeans.inertia_)

    # Wykres metody łokcia
    if show:
        plt.figure(figsize=(10, 6))
        plt.plot(range(1, 21), wss)
        plt.title('The Elbow Method')
        plt.xlabel('Number of clusters')
        plt.ylabel('WSS')
        plt.show()

    return clustering_data, scaled_data


def find_clusters_silhouette(data_tablica, show = True):
    # Usunięcie pierwszej kolumny zawierającej nazwy regionów
    clustering_data = data_tablica.drop(columns=data_tablica.columns[0])

    # Zastąpienie wszystkich '-' wartością NaN i konwersja wszystkich kolumn na typ float
    clustering_data = clustering_data.replace('-', np.nan).astype(float)

    # Wypełnienie brakujących wartości średnią dla każdej kolumny
    clustering_data = clustering_data.apply(lambda x: x.fillna(x.mean()), axis=0)

    # Normalizacja danych
    scaler = StandardScaler()
    scaled_data = scaler.fit_transform(clustering_data)

    # Obliczanie wyników silhouette dla różnej liczby klastrów
    silhouette_scores = []
    for i in range(2, 21): # silhouette score nie jest zdefiniowany dla i=1
        kmeans = KMeans(n_clusters=i, init='k-means++', max_iter=300, n_init=10, random_state=0)
        kmeans.fit(scaled_data)
        score = silhouette_score(scaled_data, kmeans.labels_)
        silhouette_scores.append(score)

    # Wykres wyników silhouette
    if show:
        plt.figure(figsize=(10, 6))
        plt.plot(range(2, 21), silhouette_scores)
        plt.title('Silhouette Scores for Different Numbers of Clusters')
        plt.xlabel('Number of clusters')
        plt.ylabel('Silhouette Score')
        plt.show()

    return scaled_data, clustering_data


def k_means(scaled_data, clustering_data, cluster_number):
    # Grupowanie K-średnich z 4 klastrami
    kmeans = KMeans(n_clusters=cluster_number, init='k-means++', max_iter=300, n_init=10, random_state=0)
    cluster_labels = kmeans.fit_predict(scaled_data)
    #print(cluster_labels)

    # Dodanie etykiet klastrów do oryginalnych danych
    clustering_data['Cluster'] = cluster_labels
    #print(clustering_data)


    # Wybór tylko kolumn numerycznych przed obliczeniem średnich wartości dla klastrów
    numerical_data = clustering_data.select_dtypes(include=[np.number])
    #print(numerical_data)

    # Obliczenie średnich wartości dla każdego klastra
    cluster_means = numerical_data.groupby('Cluster').mean()
    print(cluster_means)
    # Transpozycja danych dla łatwiejszego tworzenia wykresów
    cluster_means_transposed = cluster_means.T
    print(cluster_means_transposed)

    # Tworzenie wykresów
    plt.figure(figsize=(15, 10))
    sns.lineplot(data=cluster_means_transposed, dashes=False, markers=True)
    plt.title('Średnie Wartości Emisji Zanieczyszczeń dla Każdego Klastra')
    plt.xlabel('Zmienna')
    plt.ylabel('Średnia Emisja Zanieczyszczeń (znormalizowana)')
    plt.legend(title='Klaster', labels=cluster_means.index.tolist())
    plt.grid(True)
    plt.show()

def hierarchy(scaled_data, labels, max_d=25, p=30, truncate_mode=None):
    # Przeprowadzenie klasteryzacji hierarchicznej
    Z = linkage(scaled_data, method='ward')

    # Rysowanie dendrogramu
    plt.figure(figsize=(25, 10))
    plt.title('Hierarchical Clustering Dendrogram')
    plt.xlabel('Sample index')
    plt.ylabel('Distance')
    dendrogram(
        Z,
        leaf_rotation=90.,  # obraca etykiety liści o 90 stopni
        leaf_font_size=12.,  # rozmiar czcionki dla etykiet liści
        truncate_mode=truncate_mode,  # 'lastp' to pokazanie tylko ostatnich p połączonych klastrów
        p=p,  # p określa liczbę wyświetlanych etykiet liści
        show_contracted=True,  # pokazuje skrócony widok długich gałęzi
    )
    plt.axhline(y=max_d, color='k', linestyle='--')  # Dodanie linii poziomej na wysokości max_d
    plt.show()

    # Określenie liczby klastrów, przycinając dendrogram
    clusters = fcluster(Z, t=p, criterion='maxclust')

    # Mapowanie etykiet do klastrów
    cluster_assignment = pd.DataFrame({'Label': labels, 'Cluster': clusters})
    xlsx_filename = 'Results\\hierarchy_cluster_assignment.xlsx'
    cluster_assignment.to_excel(xlsx_filename, index=False)
    
    return clusters  # Zwrócenie przypisanych klastrów do punktów danych


# Wczytanie danych z arkusza Excel
file_path = 'Data\\Powietrze_powiaty_98.xlsx'
data_tablica = pd.read_excel(file_path, sheet_name='TABLICA')
labels = data_tablica.iloc[:, 0].values

#scaled_data, clustering_data = find_clusters_elbow(data_tablica, show=False)
#k_means(scaled_data, clustering_data, 4)

scaled_data, clustering_data = find_clusters_silhouette(data_tablica, show=False)
k_means(scaled_data, clustering_data, 2)
k_means(scaled_data, clustering_data, 3)
k_means(scaled_data, clustering_data, 4)
k_means(scaled_data, clustering_data, 5)

#scaled_data, clustering_data = find_clusters_elbow(data_tablica, show=False)
hierarchy(scaled_data, labels, truncate_mode='lastp')