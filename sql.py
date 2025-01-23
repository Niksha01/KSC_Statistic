sql = """
    -- Удаление таблиц, если они существуют
Drop TABLE IF EXISTS keys_ksc_usage;
DROP TABLE IF EXISTS institutions_ksc_keys_ksc;
DROP TABLE IF EXISTS institutions_ksc;
DROP TABLE IF EXISTS keys_ksc;
DROP TABLE IF EXISTS institutions;

-- Таблица institutions
CREATE TABLE institutions (
    institution_id SERIAL PRIMARY KEY,
    institution_name VARCHAR(255) NOT NULL
);

-- Таблица institutions_ksc
CREATE TABLE institutions_ksc (
    institution_ksc_id SERIAL PRIMARY KEY,
    institution_ksc_name VARCHAR(255) NOT NULL,
    institution_id INT,
    FOREIGN KEY (institution_id) REFERENCES institutions(institution_id)
);

-- Таблица keys_ksc
CREATE TABLE keys_ksc (
    key_ksc_id SERIAL PRIMARY KEY,
    key VARCHAR(255) NOT NULL,
    resource VARCHAR(255) NOT NULL,
    validity_period DATE NOT NULL,
    is_active  BOOLEAN DEFAULT TRUE 
);

-- Таблица institutions_ksc_keys_ksc
CREATE TABLE institutions_ksc_keys_ksc (
    institutions_ksc_keys_ksc_id SERIAL PRIMARY KEY,
    institution_ksc_id INT NOT NULL,
    key_ksc_id INT NOT NULL,
    FOREIGN KEY (institution_ksc_id) REFERENCES institutions_ksc(institution_ksc_id),
    FOREIGN KEY (key_ksc_id) REFERENCES keys_ksc(key_ksc_id)
);
-- Таблица keys_ksc_usage
CREATE TABLE keys_ksc_usage(
    keys_ksc_usage_id SERIAL PRIMARY KEY,
    key_ksc_id INT NOT NULL,
    key_usage INT NOT NULL,
    file_date DATE NOT NULL,
    FOREIGN KEY (key_ksc_id) REFERENCES keys_ksc(key_ksc_id)
);   
"""