function escapeXml(value) {
  if (value === undefined || value === null) {
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
  let current = start;
  return () => {
    current += 1;
    return String(current);
  };
}

export {
  escapeXml,
  createIdGenerator
};
