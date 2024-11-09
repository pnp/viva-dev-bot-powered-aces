using Microsoft.Bot.Builder;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Threading;

namespace WelcomeUserBotPoweredAce
{
    public class InMemoryStorage : IStorage
    {
        // Thread-safe dictionary for storing bot state in-memory
        private readonly ConcurrentDictionary<string, object> _memoryStorage = new ConcurrentDictionary<string, object>();

        /// <summary>
        /// Reads state from the in-memory store.
        /// </summary>
        /// <param name="keys">The keys of the states to read.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>A dictionary of bot state objects.</returns>
        public Task<IDictionary<string, object>> ReadAsync(string[] keys, CancellationToken cancellationToken = default)
        {
            var storeItems = new Dictionary<string, object>();

            foreach (var key in keys)
            {
                // Try to get the object from the dictionary
                if (_memoryStorage.TryGetValue(key, out var value))
                {
                    storeItems.Add(key, value);
                }
            }

            return Task.FromResult((IDictionary<string, object>)storeItems);
        }

        /// <summary>
        /// Writes state to the in-memory store.
        /// </summary>
        /// <param name="changes">Dictionary containing state objects to write.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        public Task WriteAsync(IDictionary<string, object> changes, CancellationToken cancellationToken = default)
        {
            foreach (var change in changes)
            {
                // Store the object in the dictionary
                _memoryStorage.AddOrUpdate(change.Key, change.Value, (key, oldValue) => change.Value);
            }

            return Task.CompletedTask;
        }

        /// <summary>
        /// Deletes state from the in-memory store.
        /// </summary>
        /// <param name="keys">The keys of the states to delete.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        public Task DeleteAsync(string[] keys, CancellationToken cancellationToken = default)
        {
            foreach (var key in keys)
            {
                // Remove the object from the dictionary
                _memoryStorage.TryRemove(key, out _);
            }

            return Task.CompletedTask;
        }
    }
}