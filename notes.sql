INSERT INTO pain (nome, continente, codigo)
    VALUES('Brasil','América', 'BRA'),
          ('Índia','Ásia', 'IND'),
          ('China','Ásia', 'CHI'),
          ('Japão','Ásia', 'JPN');
SELECT * FROM pais;
INSERT INTO estado (nome, sigla)
    VALUES ('Maranhão', 'MA'),
           ('São Paulo', 'SP'),
           ('Santa Catarina', 'SC'),
           ('Rio de Janeiro', 'RJ');
SELECT * FROM estado;
INSERT INTO cidade (nome, populacao)
    VALUES ('Sorocaba' 700000),
           ('Déli', 26000000),
           ('Xangai', 22000000),
           ('Tóquio', 38000000);
SELECT * FROM cidade;
INSERT INTO ponto_tur (nome, tipo)
    VALUES('Quinzinho de Barros', 'Instituição'),
          ('Parque Estadual do Jalapão', 'Atrativo'),
          ('Torre Eiffel', 'Atrativo'),
          ('Fogo de Chão', 'Restaurante');
SELECT * FROM ponto_tur; 
           
