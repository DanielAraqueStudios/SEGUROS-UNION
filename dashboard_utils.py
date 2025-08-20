import pandas as pd

def obtener_metricas(df):
    """
    Calcula métricas clave para el dashboard de seguros.
    """
    metricas = {}
    # Ejemplo: columnas esperadas
    # 'Poliza', 'Prima', 'Cliente', 'TipoSeguro', 'Fecha', 'Agente', 'Siniestro'
    if 'Prima' in df.columns:
        metricas['total_primas'] = df['Prima'].sum()
    if 'Poliza' in df.columns:
        metricas['total_polizas'] = df['Poliza'].nunique()
    if 'Cliente' in df.columns:
        metricas['total_clientes'] = df['Cliente'].nunique()
    if 'TipoSeguro' in df.columns:
        metricas['tipos_seguros'] = df['TipoSeguro'].value_counts().to_dict()
    if 'Agente' in df.columns:
        metricas['top_agentes'] = df['Agente'].value_counts().head(5).to_dict()
    if 'Siniestro' in df.columns:
        metricas['siniestros'] = df['Siniestro'].sum()
    if 'Fecha' in df.columns and 'Prima' in df.columns:
        df['Mes'] = pd.to_datetime(df['Fecha']).dt.to_period('M')
        metricas['primas_por_mes'] = df.groupby('Mes')['Prima'].sum().to_dict()
    return metricas
