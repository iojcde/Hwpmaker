function escapeXml(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function createIdGenerator(start = 2147483648) {
  let currentId = start - 1;
  return () => {
    currentId += 1;
    return currentId;
  };
}

export {
  escapeXml,
  createIdGenerator
};