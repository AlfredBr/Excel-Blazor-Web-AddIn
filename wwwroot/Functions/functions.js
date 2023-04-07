/**
 * Adds two numbers.
 * @customfunction
 * @param {number} first First number.
 * @param {number} second Second number.
 * @returns The sum of the two numbers.
 */

function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);

/**
 * Returns the volume of a sphere.
 * @customfunction
 * @param {number} radius
 */
function sphereVolume(radius) {
  return (Math.pow(radius, 3) * 4 * Math.PI) / 3;
}

CustomFunctions.associate("SPHEREVOLUME", sphereVolume);