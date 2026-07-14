class SharedStrings {
  constructor() {
    this._values = [];
    this._totalRefs = 0;
    this._hash = Object.create(null);
    this._richHash = Object.create(null);
  }

  get count() {
    return this._values.length;
  }

  get values() {
    return this._values;
  }

  get totalRefs() {
    return this._totalRefs;
  }

  getString(index) {
    return this._values[index];
  }

  add(value) {
    // Rich-text values are objects: hashing them as plain object keys would
    // stringify every one of them to '[object Object]', collapsing all
    // rich-text cells into whichever was added first. Keep a separate hash
    // keyed by the serialized value (mirrors SharedStringsXform's rich map).
    const isString = typeof value === 'string';
    const hash = isString ? this._hash : this._richHash;
    const key = isString ? value : JSON.stringify(value);
    let index = hash[key];
    if (index === undefined) {
      index = hash[key] = this._values.length;
      this._values.push(value);
    }
    this._totalRefs++;
    return index;
  }
}

module.exports = SharedStrings;
