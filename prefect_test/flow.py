from prefect import flow, task, get_run_logger

@task
def extract_data():
    logger = get_run_logger()
    data = {"message": "Hello, Prefect!"}
    logger.info(f"Extracted data: {data}")
    return data

@task
def transform_data(data):
    logger = get_run_logger()
    transformed_data = data["message"].upper()
    logger.info(f"Transformed data: {transformed_data}")
    return transformed_data

@task
def load_data(data):
    logger = get_run_logger()
    logger.info(f"Loading data: {data}")
    print(data)

@flow(name="simple-etl-flow")
def simple_etl():
    data = extract_data()
    transformed_data = transform_data(data)
    load_data(transformed_data)

if __name__ == "__main__":
    simple_etl()
