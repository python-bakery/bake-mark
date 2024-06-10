import hashlib

def hash(s):
    return int(hashlib.sha1(s.encode("utf-8")).hexdigest(), 16) % (10 ** 12)

def hash_exists(content, previous_hashes):
    return hash(content) in previous_hashes