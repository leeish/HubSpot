exports.main = async (context, sendResponse) => {
  sendResponse({
    sections: [
      {
        type: 'button',
        text: 'View Rollout',
        onClick: {
          type: 'IFRAME',
          // Width and height of the iframe (in pixels)
          width: 1500,
          height: 800,
          uri: `https://tools.hubteam.com/launch/feature-rollouts/${context.propertiesToSend.unique_hubspot_id}/info`,
        },
      },
    ],
  });
};