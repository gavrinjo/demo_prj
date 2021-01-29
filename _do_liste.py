

sample_list = [11, 31, 10, 9, 28, 1, 2]
# sample_list = sorted(sample_list)
# print(sample_list)      # sortira listu po redu od manjeg prema veÄ‡em


for i in range(len(sample_list)):
    for j in range(i + 1, len(sample_list)):
        if sample_list[i] > sample_list[j]:
            sample_list[i], sample_list[j] = sample_list[j], sample_list[i]
print(sample_list)


abc_list = ["b", "d", "a", "i", "f"]
abc_list = sorted(abc_list) # sortira lisu
abc_list.reverse()  # obrnuti redosljed
print(abc_list)


sample_list_2 = [1, 5, 18, 13, 65, 12]
print("originalan sadrÅ¾aj liste -->",  sample_list_2)
sample_list_2.sort()
print("Lista poredan od najmanjeg prema najveÄ‡em -->", sample_list_2)
sample_list_2.reverse()
print("Lista poredana od najveÄ‡eg prema najmanjem -->", sample_list_2)

