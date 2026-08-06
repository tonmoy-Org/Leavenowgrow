// Function to get all descendants (including self) of a given ID
function getDescendants(startId, data) {
  const descendants = [];
  const queue = [startId];
  const processedIds = new Set();

  while (queue.length > 0) {
    const currentId = queue.shift();
    if (processedIds.has(currentId)) {
      continue;
    }
    processedIds.add(currentId);

    const node = data.find((item) => {
      const itemId = item?.id?.toString();
      const currentDataId = currentId?.toString();

      return itemId === currentDataId;
    });

    if (node) {
      descendants.push(node);
      // Add children to the queue
      data.forEach((item) => {
        const itemPid = item?.pid?.toString();
        const currentDataId = currentId?.toString();
        if (itemPid === currentDataId && !processedIds.has(item.id)) {
          queue.push(item.id);
        }
      });
    }
  }
  return descendants;
}

function getFollowingEmployees(startId, data) {
  const descendants = [];
  const queue = [startId];
  const processedIds = new Set();

  while (queue.length > 0) {
    const currentId = queue.shift();
    if (processedIds.has(currentId)) {
      continue;
    }
    processedIds.add(currentId);

    const node = data.find((item) => {
      const itemId = item?._id?.toString();
      const currentDataId = currentId?.toString();

      return itemId === currentDataId;
    });

    if (node) {
      descendants.push(node);
      // Add children to the queue
      data.forEach((item) => {
        const itemPid = item?.manager?.id?.toString();
        const currentDataId = currentId?.toString();
        if (itemPid === currentDataId && !processedIds.has(item._id)) {
          queue.push(item._id);
        }
      });
    }
  }
  return descendants;
}

module.exports = { getDescendants, getFollowingEmployees };
