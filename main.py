
import streamlit as st
from pymongo import MongoClient
import pandas as pd
from io import BytesIO

# Connexion à MongoDB
client = MongoClient('mongodb://localhost:27017/')
db = client['word_tables']
collection = db['tables']

# Charger le CSS
st.markdown(
    """
    <style>
    .title {
        font-size: 36px;
        font-weight: bold;
        color: #4CAF50; /* Couleur verte */
    }

    .section-title {
        font-size: 24px;
        font-weight: bold;
        margin-top: 20px;
        color: #2196F3; /* Couleur bleue */
    }

    .table-container {
        border: 1px solid #ddd;
        padding: 10px;
        border-radius: 5px;
        background-color: #f9f9f9;
    }

    .table-container th {
        background-color: #4CAF50;
        color: white;
    }

    .table-container td {
        color: #333;
    }

    .separator {
        border-left: 5px solid #4CAF50;
        height: 100%;
        position: absolute;
        left: 50%;
        margin-left: -3px;
    }

    .button {
        background-color: #4CAF50;
        color: white;
        border: none;
        padding: 10px 20px;
        text-align: center;
        text-decoration: none;
        display: inline-block;
        font-size: 16px;
        margin: 4px 2px;
        cursor: pointer;
        border-radius: 5px;
    }

    .button:hover {
        background-color: #45a049;
    }

    .divider {
        height: 100%;
        width: 5px;
        background-color: #4CAF50;
        display: inline-block;
        margin: 0 10px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

def make_column_names_unique(headers):
    seen = set()
    result = []
    for item in headers:
        base_name = item
        counter = 1
        new_name = base_name
        while new_name in seen:
            new_name = f"{base_name}_{counter}"
            counter += 1
        seen.add(new_name)
        result.append(new_name)
    return result

def main():
    st.markdown('<div class="title">Sélection par titre du tableau</div>', unsafe_allow_html=True)

    # Get all regions
    all_regions = collection.distinct('region')
    selected_region = st.selectbox('Sélectionnez la région:', all_regions)

    if selected_region:
        # Get all tables for the selected region
        all_tables = list(collection.find({'region': selected_region}))

        # Extract and handle missing titles
        table_titles = []
        table_ids = []
        for table in all_tables:
            title = table.get('title', 'Table without title')  # Default title if missing
            table_titles.append(title)
            table_ids.append(table['table_id'])  # Collect table IDs for later use

        if table_titles:
            # Create a selectbox for table selection
            selected_table_title = st.selectbox('Sélectionnez le tableau:', table_titles)
            if selected_table_title:
                # Find the selected table's ID
                selected_table_id = table_ids[table_titles.index(selected_table_title)]

                # Find the selected table
                table_data = collection.find_one({'table_id': selected_table_id, 'region': selected_region})

                if table_data:
                    # Convert table data to pandas DataFrame
                    headers = make_column_names_unique(table_data['table_data'][0])
                    df = pd.DataFrame(table_data['table_data'][1:], columns=headers)  # Exclure la première ligne comme en-tête

                    # Utiliser des colonnes pour organiser l'affichage
                    col1, empty_col, col2 = st.columns([1, 0.05, 1])

                    with col1:
                        st.markdown(f'<div class="section-title">Table: {selected_table_title}</div>', unsafe_allow_html=True)
                        st.markdown('<div class="table-container">', unsafe_allow_html=True)
                        st.dataframe(df)
                        st.markdown('</div>', unsafe_allow_html=True)

                    with empty_col:
                        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

                    with col2:
                        st.markdown('<div class="section-title">Sélectionnez les variables et les lignes</div>', unsafe_allow_html=True)
                        selected_columns = st.multiselect('Sélectionnez les variables', df.columns.tolist())
                        selected_rows = st.multiselect('Sélectionnez les lignes', df.index.tolist())

                        if selected_rows or selected_columns:
                            selected_df = df.loc[selected_rows, selected_columns]
                            st.markdown('<div class="section-title">Données sélectionnées</div>', unsafe_allow_html=True)
                            st.markdown('<div class="table-container">', unsafe_allow_html=True)
                            st.table(selected_df)
                            st.markdown('</div>', unsafe_allow_html=True)

                            # Bouton pour télécharger les données sélectionnées
                            if st.button('Télécharger les données sélectionnées en Excel'):
                                output = BytesIO()
                                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                    selected_df.to_excel(writer, index=False, sheet_name='Sheet1')

                                output.seek(0)
                                st.download_button(
                                    label="Télécharger Excel",
                                    data=output,
                                    file_name='data_selection.xlsx',
                                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                                )
                else:
                    st.write("Table not found.")
        else:
            st.write("Aucun tableau trouvé pour la région sélectionnée.")

if __name__ == "__main__":
    main()
