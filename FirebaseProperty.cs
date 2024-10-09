using Google.Cloud.Firestore;

namespace Final_Project
{
    [FirestoreData]
    public class FirebaseProperty
    {
        [FirestoreProperty(nameof(ID))]
        public string ID { get; set; }

        [FirestoreProperty(nameof(PW))]
        public string PW { get; set; }

        [FirestoreProperty(nameof(Su))]
        public string Su { get; set; }

        [FirestoreProperty(nameof(Name))]
        public string Name { get; set; }
    }
}
