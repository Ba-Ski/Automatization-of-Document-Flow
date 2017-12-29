using System;
using System.Threading.Tasks;
using MongoDB.Driver;
using СurriculumParse.Structures;

namespace СurriculumParse
{
    public class MongoDbManager : IDBManager
    {
        public const string ConnectionStringName = "mongodb://localhost";
        public const string DatabaseName = "test";
        public const string CurriculumCollectionName = "Curriculum";
        public const string ErrorCollectionName = "ParseErrors";

        // This is ok... Normally, these or the entire BlogContext
        // would be put into an IoC container. We aren't using one,
        // so we'll just keep them statically here as they are 
        // thread-safe.
        private static readonly IMongoClient _client;
        private static readonly IMongoDatabase _database;

        static MongoDbManager()
        {
            _client = new MongoClient(ConnectionStringName);
            _database = _client.GetDatabase(DatabaseName);
        }

        public IMongoClient Client => _client;

        public IMongoDatabase Database => _database;

        public IMongoCollection<Curriculum> Posts => _database.GetCollection<Curriculum>(CurriculumCollectionName);

        public async Task InsertCurriculumAsync(Curriculum curriculum)
        {
            await _database.GetCollection<Curriculum>(CurriculumCollectionName).InsertOneAsync(curriculum);
        }

        public async Task ReplaceCurriculumAsync(Curriculum curriculum)
        {
            var filterBuilder = Builders<Curriculum>.Filter;
            var filter = filterBuilder.Eq(c => c.SpecialityNumber, curriculum.SpecialityNumber) &
                         filterBuilder.Eq(c => c.Profile, curriculum.Profile) &
                         filterBuilder.Eq(c => c.Year, curriculum.Year) &
                         filterBuilder.Eq(c => c.EdForm, curriculum.EdForm);
            await _database.GetCollection<Curriculum>(CurriculumCollectionName).ReplaceOneAsync(filter,curriculum,new UpdateOptions {IsUpsert = true});
        }

        public async Task<Curriculum> GetCurriculumByIdAsync(Guid id)
        {
            var filter = Builders<Curriculum>.Filter.Eq("Id", id);
            return await _database.GetCollection<Curriculum>(CurriculumCollectionName).Find(filter).FirstAsync();
        }

        public Curriculum GetCurriculumAsync(string specialityNumber, string profile, int year, int edForm)
        {
            var filterBuilder = Builders<Curriculum>.Filter;
            var filter = filterBuilder.Eq(c => c.SpecialityNumber, specialityNumber) &
                         filterBuilder.Eq(c => c.Profile, profile) &
                         filterBuilder.Eq(c => c.Year, year) &
                         filterBuilder.Eq(c => c.EdForm, edForm); 
            return _database.GetCollection<Curriculum>(CurriculumCollectionName).Find(filter).FirstOrDefault();
        }

        public async Task InsertParseError()
        {
            throw new NotImplementedException();
        }
    }
}
